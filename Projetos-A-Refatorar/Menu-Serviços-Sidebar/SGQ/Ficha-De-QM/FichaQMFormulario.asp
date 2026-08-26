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
    <script src="https://unpkg.com/just-validate@latest/dist/just-validate.production.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/select2@4.1.0-rc.0/dist/js/select2.min.js"></script>

    <script src="/corporativo/components/toast.js?v=<%=Now()%>"></script>
    <script src="/corporativo/components/confirmation.js?v=<%=Now()%>"></script>

    <script>
        let id = 0;
        let validacao;

        const listaPlantas = [
            { codigo: '00', descricao: 'Pirapetinga' },
            { codigo: '04', descricao: 'Uberaba' },
            { codigo: '05', descricao: 'Bento' },
            { codigo: '06', descricao: 'Maranguape' },
            { codigo: '07', descricao: 'Saquarema' },
            { codigo: '08', descricao: 'Araçatuba' },
            { codigo: '09', descricao: 'Blumenau' },
            { codigo: '10', descricao: 'Pacajus' },
            { codigo: '11', descricao: 'Porto Feliz' },
            { codigo: '12', descricao: 'Três Barras' },
            { codigo: '13', descricao: 'Florestal' }
        ]

        function anexarEventos() {
            $('#formularioFicha').on('submit', function(e) {
                e.preventDefault();
                salvarFormulario();
            });
        }

        function navegarParaGrade() {
            window.location.href = `FichaQMGrade.asp`;
        }

        function exibirLoading(exibir) {            
            if (exibir) {
                $('#loadingOverlay').removeClass('d-none').addClass('d-flex');
            } else {
                $('#loadingOverlay').removeClass('d-flex').addClass('d-none');
            }
        }

        function obterFicha() {
            $.ajax({
                url: 'FichaQMRepositorio.asp',
                type: 'GET',
                data: { 
                    acao: 'obterPorId',
                    id: id 
                },
                dataType: 'json',
                success: (response) => {
                    if (response.sucesso) {
                        preencherFormularioEdicao(response.dados); 
                    } else {
                        Toast.erro('Erro ao carregar ficha: ' + response.mensagem);
                    }
                },
                error: (xhr, status, error) => {
                    console.error('Erro:', error);
                    Toast.erro('Erro ao carregar dados da ficha');
                }
            });
        }

        function preencherFormularioEdicao(dados) {
            $('#idFNC').val(dados.idFNC);
            $('#dataEntradaDevolucao').val(dados.dataEntradaDevolucao);
            $('#dataConclusao').val(dados.dataConclusao);
            $('#NFDevolucao').val(dados.NFDevolucao);
            $('#codigoQM').val(dados.codigoQM);
            $('#codigoVA').val(dados.codigoVA);
            $('#quantidadePaletes').val(dados.quantidadePaletes), 
            $('#quantidadeRecebida').val(dados.quantidadeRecebida);
            $('#quantidadeRefugada').val(dados.quantidadeRefugada);
            $('#valor').val(dados.valor);
            $('#valorUnitarioMaterial').val(dados.valorUnitarioMaterial);
            $('#nomeCliente').val(dados.nomeCliente), 
            $('#modoDefeito').val(dados.modoDefeito), 
            $('#area').val(dados.area), 
            $('#ordemProducao').val(dados.ordemProducao)

            setTimeout(() => {
                $('#codPlanta').prop('disabled', true).val(dados.codPlanta).trigger('change.select2');
                $('#status').val(dados.status).trigger('change.select2');
            }, 500);
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

            exibirLoading(true)
            
            const dadosFormulario = {
                acao: id ? 'editar' : 'criar',
                codPlanta: $('#codPlanta').val(),
                status: $('#status').val(),
                idFNC: $('#idFNC').val(),
                dataEntradaDevolucao: $('#dataEntradaDevolucao').val(),
                dataConclusao: $('#dataConclusao').val(),
                NFDevolucao: $('#NFDevolucao').val(),
                codigoQM: $('#codigoQM').val(),
                codigoVA: $('#codigoVA').val(),
                quantidadePaletes: $('#quantidadePaletes').val(), 
                quantidadeRecebida: $('#quantidadeRecebida').val(),
                quantidadeRefugada: $('#quantidadeRefugada').val(),
                valor: $('#valor').val(),
                valorUnitarioMaterial: $('#valorUnitarioMaterial').val(),
                nomeCliente: $('#nomeCliente').val(), 
                modoDefeito: $('#modoDefeito').val(), 
                area: $('#area').val(), 
                ordemProducao: $('#ordemProducao').val()
            };

            if (id) dadosFormulario.id = id;

            $.ajax({
                url: 'FichaQMRepositorio.asp',
                type: 'POST',
                data: dadosFormulario,
                dataType: 'json',
                success: function(response) {
                    if (!response.sucesso) {
                        Toast.error(response.mensagem);
                        return;
                    }
                    
                    Toast.success(response.mensagem);
                    navegarParaGrade()
                },
                error: function(xhr, status, error) {
                    Toast.error('Erro ao salvar formulário');
                },
                complete: function() {
                    exibirLoading(false);
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
            $('.select2').select2({ theme: 'bootstrap-5', width: '100%', dropdownCssClass: "z-index-9999" });
        }

        function obterCamposObrigatorios(setor) {
            const camposObrigatoriosPorSetor = {
                EXPEDICAO: ['codPlanta', 'status', 'dataEntradaDevolucao', 'NFDevolucao', 'codigoVA', 'quantidadePaletes', 'quantidadeRecebida', 'quantidadeRefugada'],
                [`CUSTOMER SERVICE`]: ['valor', 'codigoQM']
            };
            return camposObrigatoriosPorSetor[setor] || [];
        }

        function atualizarCamposObrigatorios(setor) {
            const camposObrigatorios = obterCamposObrigatorios(setor);
            $('.required-field').addClass('d-none');
            camposObrigatorios.forEach(campo => {
                $(`.required-field[data-campo="${campo}"]`).removeClass('d-none');
            });
        }

        function inicializarValidacoes(setor) {
            const naoVazio = (msg) => ({
                validator: (value) => value !== null && value !== undefined && String(value).trim() !== '',
                errorMessage: msg
            });

            const maiorQueZero = (msg) => ({
                validator: (value) => value !== null && value !== undefined && parseFloat(value) > 0,
                errorMessage: msg
            })

            const regrasPorCampo = {
                codPlanta: [naoVazio('Planta é obrigatória')],
                status: [naoVazio('Status é obrigatório')],
                dataEntradaDevolucao: [naoVazio('Data de entrada devolução é obrigatória')],
                NFDevolucao: [naoVazio('NF devolução é obrigatória')],
                codigoVA: [naoVazio('Código VA é obrigatório')],
                codigoQM: [naoVazio('Código QM é obrigatório')],
                quantidadePaletes: [naoVazio('Quantidade refugada é obrigatória')],
                quantidadeRecebida: [maiorQueZero('Quantidade recebida é obrigatória')],
                quantidadeRefugada: [naoVazio('Quantidade refugada é obrigatória')],
                valor: [maiorQueZero('Valor é obrigatório')],
            };

            validacao = new JustValidate('#formularioFicha', {
                lockForm: false, 
            });

            obterCamposObrigatorios(setor).forEach(campo => {
                validacao.addField(`#${campo}`, regrasPorCampo[campo]);
            });
        }

        function obterDadosUsuario() {
            return new Promise((resolve, reject) => {
                $.ajax({
                    url: 'FichaQMRepositorio.asp?v=<%=Now()%>',
                    type: 'GET',
                    data: {
                        acao: 'obterDadosUsuario',
                    },
                    dataType: 'json',
                    success: function(response) {
                        exibirLoading(false);
                        resolve(response); 
                    },
                    error: function(xhr, status, error) {
                        exibirLoading(false);
                        reject(error); 
                    }
                });
            });
        }

        function popularPlantas(plantasString) {
            const plantasUsuario = plantasString.split(',').map(planta => planta.trim());
            
            const $select = $('#codPlanta');
            $select.empty().append('<option value="">Escolha uma planta</option>');

            const plantasVisiveis = listaPlantas.filter(planta => plantasUsuario.includes(planta.codigo));
            plantasVisiveis.forEach(function(item) {
                $select.append($('<option>', { value: item.codigo, text: item.descricao }));
            });
        }

        function carregarStatus() {
            const $select = $('#status');
            $select.empty().append('<option value="">Selecione o status</option>');

            const params = {
                acao: 'listarStatus'
            };

            $.ajax({
                url: 'FichaQMRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: params,
                dataType: 'json',
                success: function(response) {
                    if (!response.sucesso) {
                        Toast.error("Erro ao carregar status: " + response.mensagem);
                        return;
                    }

                    response.dados.forEach(function (item) {
                        $select.append($('<option>', { value: item.codigo, text: item.descricao }));
                    });
                },
                error: function(xhr, status, error) {
                    console.error('Erro ao carregar status:', error);
                }
            });
        }


        $(document).ready(async function() {
            const urlParams = new URLSearchParams(window.location.search);
            const paramId = urlParams.get('id');

            if (paramId) {
                id = paramId;
                $('#tituloPagina').text(`Editar Ficha de QM #${id}`);
                $('#codPlanta').prop('disabled', true);
                obterFicha();
            } else {
                $('#tituloPagina').text('Nova Ficha de QM');
            }

            resposta = await obterDadosUsuario();
            if (!resposta.sucesso) {
                alert("Você não tem permissão para acessar esta página.");
                window.location.href = '/corporativo/frameset/principal.asp';
                return;
            }
        
            const dadosUsuario = resposta.dados;
            if (!dadosUsuario || dadosUsuario.nivel != 'RESPONSAVEL') {
                alert("Você não tem permissão para acessar esta página.");
                window.location.href = '/corporativo/frameset/principal.asp';
                return;
            }

            if(!['EXPEDICAO', 'CUSTOMER SERVICE'].includes(dadosUsuario.setor)) {
                alert("Você não tem permissão para acessar esta página.");
                window.location.href = '/corporativo/frameset/principal.asp';
                return;
            }

            if(dadosUsuario.setor == 'EXPEDICAO') { 
                // Expedição edita tudo menos valor e código QM
                $('#valor').prop('disabled', true);
                $('#codigoQM').prop('disabled', true);
            } else {
                // Controle de qualidade edita apenas valor e codigo QM
                $('#idFNC').prop('disabled', true);
                $('#codPlanta').prop('disabled', true);
                $('#status').prop('disabled', true);
                $('#dataEntradaDevolucao').prop('disabled', true);
                $('#dataConclusao').prop('disabled', true);
                $('#NFDevolucao').prop('disabled', true);
                $('#codigoVA').prop('disabled', true);
                $('#quantidadeRecebida').prop('disabled', true);
                $('#quantidadeRefugada').prop('disabled', true);
                $('#quantidadePaletes').prop('disabled', true);
                $('#nomeCliente').prop('disabled', true);
                $('#modoDefeito').prop('disabled', true);
                $('#area').prop('disabled', true);
                $('#ordemProducao').prop('disabled', true);
                $('#quantidadeRefugada').prop('disabled', true);
                $('#valorUnitarioMaterial').prop('disabled', true);
            }

            carregarStatus();
            popularPlantas(dadosUsuario.plantas);
            atualizarCamposObrigatorios(dadosUsuario.setor);
            inicializarValidacoes(dadosUsuario.setor);
            anexarEventos();
            inicializarSelect2();
        });
    </script>
</head>
<body>
    <div class="container my-4">
        <div class="d-flex my-4">
            <button onclick="navegarParaGrade()" class="btn btn-sm btn-light me-2 fs-5"> 
                <i class="bi bi-arrow-left"></i> 
            </button>
            <h3 id="tituloPagina"></h3>
        </div>

        <div id="loadingOverlay" class="position-fixed top-0 start-0 w-100 h-100 d-none justify-content-center align-items-center" style="background: rgba(255,255,255,0.8); z-index: 9999;">
            <div class="spinner-border text-primary" role="status">
                <span class="visually-hidden">Carregando...</span>
            </div>
        </div>

        <div class="card p-3">
            <form id="formularioFicha" name="formularioFicha" novalidate>
                <div class="row g-3">
                    <div class="col-12">
                        <label for="codPlanta" class="form-label">Planta <span class="required-field" data-campo="codPlanta">*</span></label>
                        <select class="form-select form-select-sm select2" name="codPlanta" id="codPlanta">
                            <option value="">Selecione a planta</option>
                        </select>
                    </div>
                    <div class="col-12 col-md-4">
                        <label for="idFNC" class="form-label">FNC </label>
                        <input type="number" class="form-control form-control-sm" id="idFNC" name="idFNC">
                    </div>
                    <div class="col-12 col-md-4">
                        <label for="status" class="form-label">Status <span class="required-field" data-campo="status">*</span></label>
                        <select class="form-select form-select-sm select2" name="status" id="status">
                            <option value="">Selecione o status</option>
                        </select>
                    </div>
                    <div class="col-12 col-md-4">
                        <label for="nomeCliente" class="form-label">Cliente</label>
                        <input type="text" class="form-control form-control-sm" id="nomeCliente" name="nomeCliente">
                    </div>
                    <div class="col-12 col-md-4">
                        <label for="modoDefeito" class="form-label">Modo de Defeito</label>
                        <input type="text" class="form-control form-control-sm" id="modoDefeito" name="modoDefeito">
                    </div>
                    <div class="col-12 col-md-4">
                        <label for="area" class="form-label">Área</label>
                        <input type="text" class="form-control form-control-sm" id="area" name="area">
                    </div>
                    <div class="col-12 col-md-4">
                        <label for="ordemProducao" class="form-label">Ordem de produção</label>
                        <input type="text" class="form-control form-control-sm" id="ordemProducao" name="ordemProducao">
                    </div>
                    <div class="col-12 col-md-4">
                        <label for="dataEntradaDevolucao" class="form-label">Data Entrada Devolução <span class="required-field" data-campo="dataEntradaDevolucao">*</span> </label>
                        <input type="date" class="form-control form-control-sm" id="dataEntradaDevolucao" name="dataEntradaDevolucao">
                    </div>
                     <div class="col-12 col-md-4">
                        <label for="dataConclusao" class="form-label">Data de Conclusão <span class="required-field" data-campo="dataConclusao">*</span> </label>
                        <input type="date" class="form-control form-control-sm" id="dataConclusao" name="dataConclusao">
                    </div>
                    <div class="col-12 col-md-4">
                        <label for="NFDevolucao" class="form-label">NF Devolução <span class="required-field" data-campo="NFDevolucao">*</span> </label>
                        <input type="text" class="form-control form-control-sm" id="NFDevolucao" name="NFDevolucao">
                    </div>
                    <div class="col-12 col-md-4">
                        <label for="codigoVA" class="form-label">Código VA <span class="required-field" data-campo="codigoVA">*</span> </label>
                        <input type="text" class="form-control form-control-sm" id="codigoVA" name="codigoVA">
                    </div>
                    <div class="col-12 col-md-4">
                        <label for="quantidadePaletes" class="form-label">Quantidade Paletetes <span class="required-field" data-campo="quantidadePaletes">*</span> </label>
                        <input value="0" type="number" class="form-control form-control-sm" id="quantidadePaletes" name="quantidadePaletes">
                    </div>
                    <div class="col-12 col-md-4">
                        <label for="quantidadeRecebida" class="form-label">Quantidade Recebida <span class="required-field" data-campo="quantidadeRecebida">*</span> </label>
                        <input value="0" type="number" class="form-control form-control-sm" id="quantidadeRecebida" name="quantidadeRecebida">
                    </div>
                    <div class="col-12 col-md-4">
                        <label for="quantidadeRefugada" class="form-label">Quantidade Refugada <span class="required-field" data-campo="quantidadeRefugada">*</span> </label>
                        <input value="0" type="number" class="form-control form-control-sm" id="quantidadeRefugada" name="quantidadeRefugada">
                    </div>
                    <div class="col-12 col-md-4">
                        <label for="codigoQM" class="form-label">Código QM <span class="required-field" data-campo="codigoQM">*</span> </label>
                        <input type="text" class="form-control form-control-sm" id="codigoQM" name="codigoQM">
                    </div>
                    <div class="col-12 col-md-4">
                        <label for="valor" class="form-label">Valor <span class="required-field" data-campo="valor">*</span> </label>
                        <input value="0" type="number" class="form-control form-control-sm" id="valor" name="valor">
                    </div>
                    <div class="col-12 col-md-4">
                        <label for="valorUnitarioMaterial" class="form-label">Valor Unit. do Material </label>
                        <input value="0" type="number" class="form-control form-control-sm" id="valorUnitarioMaterial" name="valorUnitarioMaterial">
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
