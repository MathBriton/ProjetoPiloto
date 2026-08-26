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
    <title>Lote Piloto</title>

    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.1/font/bootstrap-icons.css"/>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.0.2/dist/css/bootstrap.min.css" rel="stylesheet" integrity="sha384-EVSTQN3/azprG1Anm3QDgpJLIm9Nao0Yz1ztcQTwFspd3yD65VohhpuuCOmLASjC" crossorigin="anonymous"/>
    <link href="https://cdn.jsdelivr.net/npm/select2@4.1.0-rc.0/dist/css/select2.min.css" rel="stylesheet"/>
    <link href="https://cdn.jsdelivr.net/npm/select2-bootstrap-5-theme@1.3.0/dist/select2-bootstrap-5-theme.min.css" rel="stylesheet"/>

    <style>
        body {
            background-color: #f5f5f5;
            font-size: .9rem;
        }
        .container {
            max-width: 85%;
        }
        .page-actions {
            gap: .5rem;
        }
        .form-content {
            padding: 30px;
        }
        .form-label {
            font-weight: 600;
        }
        .required-field {
            color: #dc3545;
        }
        .badge {
            font-size: .85rem !important;
        }
        .section-header {
            border-bottom: 1px solid #e9ecef;
            margin-bottom: 20px;
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
            font-weight: 500;
            color: #495057;
            padding-left: 8px;
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
        .z-index-9999 {
            z-index: 9999 !important;
        }
        @media (max-width: 1400px) {
            .container {
                max-width: 95%;
            }
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
        .badge {
            font-size: .75rem;
        }
        .badge-red {
            background-color: #fee2e2;
            color: #991b1b;
        }
        .badge-green {
            background-color: #d1fae5;
            color: #065f46;
        }      
        .badge-yellow {
            background-color: #fef3c7;
            color: #92400e;
        }  
    </style>

    <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.0.2/dist/js/bootstrap.bundle.min.js" integrity="sha384-MrcW6ZMFYlzcLA8Nl+NtUVF0sA7MsXsP1UyJoMp4YLEuNSfAP+JcXn/tWtIaxVXM" crossorigin="anonymous"></script>
    <script src="https://cdn.jsdelivr.net/npm/select2@4.1.0-rc.0/dist/js/select2.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/just-validate@latest/dist/just-validate.production.min.js"></script>
    <script src="/corporativo/components/toast.js?v=<%=Now()%>"></script>
    <script src="/corporativo/components/confirmation.js?v=<%=Now()%>"></script>
    <script src="/corporativo/funcoes/funcoesJAVASCRIPT.js?v=<%=Now()%>"></script>

    <script>
        let id = null;
        let acaoFormulario = 'CRIAR';
        let tipoContinuacao = null;
        let codigoProdutoAnterior = null;
        let validacao = null;
        let statusAtual = null;
        let valorLotePilotoAtual = null;
        let quantidadePedidosAtual = null;

        const plantasWestrock = ['08', '09', '10', '11', '12', '13'];

        // Lista de status que usuário não pode selecionar (são status automaticos, setados em outras rotinas)
        const statusAutomaticos = ['1', '2', '7', '9'];

        // Lista de status que necessiram de justificativa
        const statusComJustificativaObrigatoria = ['3', '5', '6', '10'];

        // Busca de Cliente
        let clienteSelecionado = null;
        let resultadosCliente = [];

        function exibirLoading(exibir) {
            if (exibir) {
                $('#loadingOverlay').removeClass('d-none').addClass('d-flex');
            } else {
                $('#loadingOverlay').removeClass('d-flex').addClass('d-none');
            }
        }

        function inicializarSelect2() {
            $('.select2').select2({
                theme: 'bootstrap-5',
                width: '100%',
                dropdownCssClass: 'z-index-9999'
            });
        }

        function carregarTiposItem() {
            const $select = $('#tipoItem');
            $select.empty().append('<option value="">Selecione um item</option>');
            return $.ajax({
                url: 'LotePilotoRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: { acao: 'listarTiposItem' },
                dataType: 'json',
                success: function(response) {
                    if (!response.sucesso) {
                        console.error('Erro ao carregar tipos de item: ' + response.mensagem);
                        return;
                    }
                    (response.dados || []).forEach(function(item) {
                        $select.append($('<option>', { value: item.codigo, text: item.descricao }));
                    });
                    $select.trigger('change.select2');
                },
                error: function(xhr, status, error) {
                    console.error('Erro ao carregar tipos de item:', error);
                }
            });
        }

        function carregarMotivos() {
            const $select = $('#motivo');
            $select.empty().append('<option value="">Selecione um item</option>');
            return $.ajax({
                url: 'LotePilotoRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: { acao: 'listarMotivos' },
                dataType: 'json',
                success: function(response) {
                    if (!response.sucesso) {
                        console.error('Erro ao carregar motivos: ' + response.mensagem);
                        return;
                    }
                    (response.dados || []).forEach(function(item) {
                        $select.append($('<option>', { value: item.codigo, text: item.descricao }));
                    });
                    $select.trigger('change.select2');
                },
                error: function(xhr, status, error) {
                    console.error('Erro ao carregar motivos:', error);
                }
            });
        }


        function carregarTiposLotePiloto() {
            const $select = $('#tipoLP');
            $select.empty().append('<option value="">Selecione um item</option>');
            return $.ajax({
                url: 'LotePilotoRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: { acao: 'listarTiposLotePiloto' },
                dataType: 'json',
                success: function(response) {
                    if (!response.sucesso) {
                        console.error('Erro ao carregar tipos de lote piloto: ' + response.mensagem);
                        return;
                    }
                    (response.dados || []).forEach(function(item) {
                        $select.append($('<option>', { value: item.codigo, text: item.descricao }));
                    });
                    $select.trigger('change.select2');
                },
                error: function(xhr, status, error) {
                    console.error('Erro ao carregar tipos de lote piloto:', error);
                }
            });
        }

        function carregarStatus() {
            const $select = $('#status');
            $select.empty().append('<option value="">Selecione um item</option>');

            return $.ajax({
                url: 'LotePilotoRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: { acao: 'listarStatus' },
                dataType: 'json',
                success: function(response) {
                    if (!response.sucesso) {
                        console.error('Erro ao carregar status: ' + response.mensagem);
                        return;
                    }

                    (response.dados || []).forEach(function(item) {
                        $select.append($('<option>', { value: item.codigo, text: item.descricao }));
                    });

                    $select.trigger('change.select2');
                },
                error: function(xhr, status, error) {
                    console.error('Erro ao carregar status:', error);
                }
            });
        }

        function carregarDadosLotePiloto() {
            return $.ajax({
                url: 'LotePilotoRepositorio.asp',
                type: 'GET',
                data: { acao: 'obterPorId', id: id },
                dataType: 'json',
                success: function(response) {
                    if (!response.sucesso) {
                        Toast.error(response.mensagem);
                        return;
                    }
                    preencherFormulario(response.dados);
                },
                error: function(xhr, status, error) {
                    console.error('Erro ao carregar dados:', error);
                    Toast.error('Erro ao carregar dados');
                }
            });
        }

        function preencherDadosContinuacao(dados) {
            if (!dados) return;

            codigoProdutoAnterior = dados.codigoProduto;
            $('#codigoPlanta').val(dados.codigoPlanta).trigger('change.select2').trigger('change');
            $('#codigoCliente').val(dados.codigoCliente);
            $('#codigoLoja').val(dados.codigoLoja);
            $('#nomeCliente').val(dados.nomeCliente);

            // MESMO_PRODUTO também copia os dados do produto anterior.
            if (tipoContinuacao === 'MESMO_PRODUTO') {
                $('#tipoItem').val(dados.tipoItem).trigger('change.select2');
                $('#codigoProduto').val(dados.codigoProduto);
                $('#referencia').val(dados.referencia);
            }
        }

        function preencherFormulario(dados) {
            if (!dados) return;

            if (acaoFormulario === 'CONTINUAR') {
                preencherDadosContinuacao(dados);
                return;
            }

            $('#codigoPlanta').val(dados.codigoPlanta).trigger('change.select2').trigger('change');
            $('#codigoCliente').val(dados.codigoCliente);
            $('#codigoLoja').val(dados.codigoLoja);
            $('#nomeCliente').val(dados.nomeCliente);
            $('#tipoItem').val(dados.tipoItem).trigger('change.select2');
            $('#codigoProduto').val(dados.codigoProduto);
            $('#referencia').val(dados.referencia);
            $('#quantidadePaletes').val(dados.quantidadePaletes);
            $('#pedidoWRK').val(dados.pedidoWRK);
            $('#motivo').val(dados.motivo).trigger('change.select2');
            $('#precisaLiberacao').val(dados.precisaLiberacao).trigger('change.select2');
            $('#tipoLP').val(dados.tipoLP).trigger('change.select2');

            valorLotePilotoAtual = dados.lotePilotoAprovado
            $('#lotePilotoAprovado').val(dados.lotePilotoAprovado).trigger('change.select2');

            statusAtual = String(dados.status);
            $('#status').val(statusAtual).trigger('change.select2');

            quantidadePedidosAtual = Number(dados.quantidadePedidos || 0);
            $('#quantidadePedidos').val(dados.quantidadePedidos);

            atualizarJustificativas();

            $('#acompanhamentoOnduladeira').prop('checked', dados.acompanhamentoOnduladeira == '1');
            $('#acompanhamentoImpressora').prop('checked', dados.acompanhamentoImpressora == '1');
            $('#acompanhamentoExpedicao').prop('checked', dados.acompanhamentoExpedicao == '1');
            $('#acompanhamentoCliente').prop('checked', dados.acompanhamentoCliente == '1');
            $('#acompanhamentoDDP').prop('checked', dados.acompanhamentoDDP == '1');
            $('#acompanhamentoClicheria').prop('checked', dados.acompanhamentoClicheria == '1');
        }

        function inicializarValidacao() {
            validacao = new JustValidate('#formularioUnificado', { lockForm: false });
            validacao
                .addField('#codigoCliente', [{ rule: 'required', errorMessage: 'Código do cliente é obrigatório' }])
                .addField('#codigoLoja', [{ rule: 'required', errorMessage: 'Loja é obrigatória' }])
                .addField('#nomeCliente', [{ rule: 'required', errorMessage: 'Nome do cliente é obrigatório' }])
                .addField('#codigoProduto', [{ rule: 'required', errorMessage: 'Produto é obrigatório' }])
                .addField('#quantidadePaletes', [{ rule: 'required', errorMessage: 'Quantidade de palete é obrigatória' }])
                .addField('#quantidadePedidos', [{ rule: 'required', errorMessage: 'Quantidade de pedidos é obrigatória' }])
                .addField('#motivo', [{ rule: 'required', errorMessage: 'Motivo é obrigatório' }])
                .addField('#precisaLiberacao', [{ rule: 'required', errorMessage: 'Liberação é obrigatória' }])
                .addField('#tipoLP', [{ rule: 'required', errorMessage: 'Tipo de LP é obrigatório' }])
                .addField('#pedidoWRK', [{
                    validator: function(value) {
                        return !plantasWestrock.includes(String($('#codigoPlanta').val())) || String(value || '').trim() !== '';
                    },
                    errorMessage: 'Pedido KIWI é obrigatório para plantas WestRock'
                }]);

            if (acaoFormulario !== 'EDITAR') {
                validacao.addField('#tipoItem', [{ rule: 'required', errorMessage: 'Tipo do item é obrigatório' }]);
            }

            if (acaoFormulario === 'EDITAR') {
                validacao
                    .addField('#lotePilotoAprovado', [{ rule: 'required', errorMessage: 'Aprovação é obrigatória' }])
                    .addField('#status', [{ rule: 'required', errorMessage: 'Status é obrigatório' }]);
            }
        }

        function obterJustificativasObrigatorias() {
            const justificativas = [];
            const quantidadePedidos = Number($('#quantidadePedidos').val() || 0);
            const novoStatus = String($('#status').val() || '');

            const quantidadePedidosAlterada = acaoFormulario !== 'EDITAR' || quantidadePedidos !== quantidadePedidosAtual;

            if (quantidadePedidosAlterada && quantidadePedidos > 2) {
                justificativas.push({ tipoJustificativa: 1, motivo: 'Quantidade de pedidos LP maior que 2' });
            }

            if (acaoFormulario === 'EDITAR' && novoStatus !== statusAtual && statusComJustificativaObrigatoria.includes(novoStatus)) {
                justificativas.push({
                    tipoJustificativa: novoStatus,
                    motivo: $('#status option:selected').text()
                });
            }

            return justificativas;
        }

        function atualizarJustificativas() {
            const justificativas = obterJustificativasObrigatorias();
            const $container = $('#camposJustificativa');
            const valoresAtuais = {};

            $container.find('[data-tipo-justificativa]').each(function() {
                valoresAtuais[$(this).data('tipo-justificativa')] = $(this).val();
            });

            $container.empty();

            justificativas.forEach(function(justificativa) {
                const tipo = justificativa.tipoJustificativa;
                const valor = valoresAtuais[tipo] || '';
                $container.append(`
                    <div class="col-12">
                        <label for="descricaoJustificativa${tipo}" class="form-label">
                            Justificativa: ${justificativa.motivo} <span class="required-field">*</span>
                        </label>
                        <textarea class="form-control form-control-sm descricao-justificativa" rows="3"
                            id="descricaoJustificativa${tipo}" data-tipo-justificativa="${tipo}">${valor}</textarea>
                    </div>
                `);
            });
        }

        function validarJustificativas() {
            let justificativasValidas = true;

            $('.descricao-justificativa').each(function() {
                const preenchida = String($(this).val() || '').trim() !== '';
                $(this).toggleClass('is-invalid', !preenchida);
                justificativasValidas = justificativasValidas && preenchida;
            });

            if (!justificativasValidas) {
                Toast.error('Preencha todas as justificativas obrigatórias.');
            }

            return justificativasValidas;
        }

        async function validarFormulario() {
            atualizarJustificativas();

            const formularioValido = await validacao.revalidate();
            if (!formularioValido) {
                exibirErrosValidacao(validacao);
                return;
            }

            if (statusAutomaticos.includes($('#status').val()) && String($('#status').val()) !== statusAtual) {
                Toast.error('Esse status não pode ser adicionado manualmente')
                return; 
            }
            
            if (!validarJustificativas()) return;
            enviarFormulario();
        }

        function enviarFormulario() {
            const dadosFormulario = {
                codigoPlanta: $('#codigoPlanta').val(),
                codigoCliente: $('#codigoCliente').val(),
                codigoLoja: $('#codigoLoja').val(),
                nomeCliente: $('#nomeCliente').val(),
                tipoItem: $('#tipoItem').val(),
                lotePilotoAprovado: $('#lotePilotoAprovado').val(),
                status: $('#status').val(),
                statusAnterior: statusAtual,
                codigoProduto: $('#codigoProduto').val(),
                referencia: $('#referencia').val(),
                quantidadePaletes: $('#quantidadePaletes').val(),
                quantidadePedidos: $('#quantidadePedidos').val(),
                pedidoWRK: $('#pedidoWRK').val(),
                motivo: $('#motivo').val(),
                motivoDescricao: $("#motivo option:selected").text(),
                precisaLiberacao: $('#precisaLiberacao').val(),
                tipoLP: $('#tipoLP').val(),
                justificativas: JSON.stringify($('.descricao-justificativa').map(function() {
                    return {
                        tipoJustificativa: $(this).data('tipo-justificativa'),
                        descricaoJustificativa: String($(this).val() || '').trim()
                    };
                }).get()),
                acompanhamentoOnduladeira: $('#acompanhamentoOnduladeira').is(':checked') ? '1' : '0',
                acompanhamentoImpressora: $('#acompanhamentoImpressora').is(':checked') ? '1' : '0',
                acompanhamentoExpedicao: $('#acompanhamentoExpedicao').is(':checked') ? '1' : '0',
                acompanhamentoCliente: $('#acompanhamentoCliente').is(':checked') ? '1' : '0',
                acompanhamentoDDP: $('#acompanhamentoDDP').is(':checked') ? '1' : '0',
                acompanhamentoClicheria: $('#acompanhamentoClicheria').is(':checked') ? '1' : '0',
                aprovandoLotePiloto: $('#lotePilotoAprovado').val() == '1' && (valorLotePilotoAtual != $('#lotePilotoAprovado').val()) ? '1' : '0'
            };

            if (acaoFormulario === 'EDITAR') {
                dadosFormulario.id = id;
                dadosFormulario.acao = 'editar';
            } else {
                dadosFormulario.acao = 'criar';
                if (acaoFormulario === 'CONTINUAR') {
                    dadosFormulario.codigoProdutoAnterior = codigoProdutoAnterior;
                }
            }

            exibirLoading(true);

            $.ajax({
                url: 'LotePilotoRepositorio.asp',
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
                    Toast.error('Erro ao salvar formulário');
                },
                complete: function() {
                    exibirLoading(false);
                }
            });
        }

        function navegarParaGrade() {
            window.location.href = 'LotePilotoGrade.asp';
        }

        function abrirModalCliente() {
            const codigoPlanta = $('#codigoPlanta').val();

            if (!codigoPlanta) { 
                Toast.warning('Informe o código da planta antes de buscar por clientes.'); 
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
                        <p class="mt-2 mb-0">Digite um termo para buscar cliente existentes.</p>
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
                url: 'LotePilotoRepositorio.asp?v=<%=Now()%>',
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
                                <p class="mt-2 mb-0"> ${response.mensagem || 'Erro ao buscar clientes.'} </p>
                            </div>
                        `);
                        return;
                    }

                    resultadosCliente = response.dados || [];

                    if (resultadosCliente.length === 0) {
                        $('#clienteResultados').html(`
                            <div class="alert text-center alert-info" role="alert">
                                <i class="bi bi-info-circle" style="font-size: 1.5rem;"></i>
                                <p class="mt-2 mb-0"> Nenhum cliente encontrado com os termos pesquisados. </p>
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
                            <p class="mt-2 mb-0"> Nenhum cliente encontrado com os termos pesquisados.</p>
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
                                <th>Bairro</th>
                                <th>Municipio</th>
                                <th>UF</th>
                            </tr>
                        </thead>
                        <tbody>
            `;

            resultadosCliente.forEach(function(cliente, index) {
                html += `
                    <tr data-index="${index}">
                        <td class="text-center">
                            <button onclick="selecionarCliente(${index})"
                                class="btn-selecionar" 
                                title="Selecionar este cliente">
                                <i class="bi bi-circle"></i> 
                            </button>
                        </td>
                        <td>${cliente.nome}</td>
                        <td>${cliente.codigoCliente}</td>
                        <td>${cliente.codigoLoja}</td>
                        <td>${formatarCNPJ(cliente.cnpj)}</td>
                        <td>${cliente.bairro}</td>
                        <td>${cliente.municipio}</td>
                        <td>${cliente.UF}</td>
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
        }

        function preencherDadosCliente(dados) {
            $('#nomeCliente').val(dados.nome || '');
            $('#codigoCliente').val(dados.codigoCliente || '');
            $('#codigoLoja').val(dados.codigoLoja || '');
            $('#cnpj').val(dados.cnpj || '').trigger('input');
        }

        /* Aba de Produtos */
        function abrirModalProduto() {
            const codigoPlanta = $('#codigoPlanta').val();
            const codigoCliente = $('#codigoCliente').val();

            if (!codigoPlanta) { 
                Toast.warning('Informe o código da planta antes de buscar por produtos.'); 
                return; 
            }

            if (!codigoCliente) {
                Toast.warning('Informe o código do cliente antes de buscar por produtos.');
                return;
            }

            produtoSelecionado = null;
            resultadosProduto = [];

            $('#produtoBusca').val('');
            $('#produtoResultados').html(`
                <div class="text-center text-muted">
                    <i class="bi bi-search" style="font-size: 1.5rem;"></i>
                    <p class="mt-2 mb-0">Digite um termo para buscar produtos existentes.</p>
                </div>
            `);
            $('#btnConfirmarProduto').prop('disabled', true);

            const modal = new bootstrap.Modal(document.getElementById('modalPesquisaProduto'));
            modal.show();

            setTimeout(function() { 
                $('#produtoBusca').focus(); 
            }, 500);
        }

        function pesquisarProduto() {
            const busca = $('#produtoBusca').val().trim();
            const tipoItem = $('#tipoItem').val();
            const codigoPlanta = $('#codigoPlanta').val();
            const codigoCliente = $('#codigoCliente').val();
            const codigoLoja = $('#codigoLoja').val();

            if (!busca || busca.length < 2) {
                $('#produtoResultados').html(`
                    <div class="text-center text-muted">
                        <i class="bi bi-search" style="font-size: 1.5rem;"></i>
                        <p class="mt-2 mb-0">Digite um termo para buscar produtos existentes.</p>
                    </div>
                `);
                return;
            }

            produtoSelecionado = null;
            $('#btnConfirmarProduto').prop('disabled', true);

            $('#produtoResultados').html(`
                <div class="text-center">
                    <div class="spinner-border spinner-border-sm text-primary" role="status">
                        <span class="visually-hidden">Buscando...</span>
                    </div>
                    <p class="mt-2 mb-0 text-muted">Buscando produtos...</p>
                </div>
            `);

            $.ajax({
                url: 'LotePilotoRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: {
                    acao: 'listarProdutos',
                    tipoItem: tipoItem,
                    codigoPlanta: codigoPlanta,
                    codigoCliente: codigoCliente,
                    codigoLoja: codigoLoja,
                    busca: busca
                },
                dataType: 'json',
                success: function(response) {
                    if (!response.sucesso) {
                        $('#produtoResultados').html(`
                            <div class="alert text-center alert-danger" role="alert">
                                <i class="bi bi-exclamation-triangle" style="font-size: 1.5rem;"></i>
                                <p class="mt-2 mb-0"> ${response.mensagem || 'Erro ao buscar produtos.'} </p>
                            </div>
                        `);
                        return;
                    }

                    resultadosProduto = response.dados || [];

                    if (resultadosProduto.length === 0) {
                        $('#produtoResultados').html(`
                            <div class="alert text-center alert-info" role="alert">
                                <i class="bi bi-info-circle" style="font-size: 1.5rem;"></i>
                                <p class="mt-2 mb-0"> Nenhum produto encontrado com os termos pesquisados. </p>
                            </div>
                        `);
                        return;
                    }

                    renderizarResultadosProduto();
                },
                error: function(xhr, status, error) {
                    console.error('Erro ao buscar produtos:', error);
                    $('#produtoResultados').html(`
                        <div class="alert text-center alert-info" role="alert">
                            <i class="bi bi-info-circle" style="font-size: 1.5rem;"></i>
                            <p class="mt-2 mb-0"> Nenhum produto encontrado com os termos pesquisados.</p>
                        </div>
                    `);
                }
            });
        }

        function renderizarResultadosProduto() {
            let html = `
                <div class="table-responsive">
                    <table class="table table-sm table-hover mb-0">
                        <thead class="thead-light">
                            <tr>
                                <th style="width: 50px;"></th>
                                <th>Código</th>
                                <th>Referência</th>
                            </tr>
                        </thead>
                        <tbody>
            `;

            resultadosProduto.forEach(function(produto, index) {
                html += `
                    <tr data-index="${index}">
                        <td class="text-center">
                            <button onclick="selecionarProduto(${index})"
                                class="btn-selecionar" 
                                title="Selecionar este produto">
                                <i class="bi bi-circle"></i> 
                            </button>
                        </td>
                        <td>${produto.codigo}</td>
                        <td>${produto.referencia}</td>
                    </tr>
                `;
            });

            html += `
                        </tbody>
                    </table>
                </div>
            `;

            $('#produtoResultados').html(html);
        }

        function selecionarProduto(index) {
            produtoSelecionado = resultadosProduto[index];
            
            $('#produtoResultados .btn-selecionar')
                .removeClass('selecionado')
                .html('<i class="bi bi-circle"></i>')
                .attr('title', 'Selecionar este item');
            
            $('#produtoResultados tr[data-index="' + index + '"] .btn-selecionar')
                .addClass('selecionado')
                .html('<i class="bi bi-check-circle-fill"></i>')
                .attr('title', 'Item Selecionado');
            
            $('#btnConfirmarProduto').prop('disabled', false);
        }

        function confirmarProduto() {
            if (!produtoSelecionado) {
                Toast.warning('Selecione um produto');
                return;
            }

            const modal = bootstrap.Modal.getInstance(document.getElementById('modalPesquisaProduto'));
            modal.hide();

            preencherDadosProduto(produtoSelecionado);
        }

        function preencherDadosProduto(dados) {
            $('#codigoProduto').val(dados.codigo);
            $('#referencia').val(dados.referencia);
        }

        function anexarEventos() {
            $('#btnSalvar').on('click', function() {
                validarFormulario();
            });

            $('#formularioUnificado').on('submit', function(e) {
                e.preventDefault();
                validarFormulario();
            });

            $('#quantidadePedidos, #status').on('change input', function() {
                atualizarJustificativas();
            });

            $('#codigoPlanta').on('change', function() {
                const pedidoKiwiObrigatorio = plantasWestrock.includes(String($(this).val()));
                $('#grupoPedidoKiwi').toggleClass('d-none', !pedidoKiwiObrigatorio);
                
                if (!pedidoKiwiObrigatorio) $('#pedidoWRK').val('');
            });

            /* Busca de Cliente */
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

             /* Busca de Produto */
            let buscaProdutoTimeout = null;

            $('#produtoBusca').on('input', function() {
                clearTimeout(buscaProdutoTimeout);
                buscaProdutoTimeout = setTimeout(function() {
                    pesquisarProduto();
                }, 500);
            });

            $('#produtoBusca').on('keypress', function(e) {
                if (e.which === 13) {
                    e.preventDefault();
                    clearTimeout(buscaProdutoTimeout);
                    pesquisarProduto();
                }
            });
        }

        $(document).ready(async function() {
            const urlParams = new URLSearchParams(window.location.search);
            acaoFormulario = String(urlParams.get('acao') || 'CRIAR').toUpperCase();
            tipoContinuacao = String(urlParams.get('tipo') || '').toUpperCase();
            id = urlParams.get('id');

            if (!['CRIAR', 'EDITAR', 'CONTINUAR'].includes(acaoFormulario)) {
                acaoFormulario = 'CRIAR';
                id = null;
            }
            
            if (acaoFormulario === 'CRIAR') id = null;

            exibirLoading(true);

            try {
                anexarEventos();
                inicializarSelect2();

                await Promise.all([
                    carregarTiposItem(),
                    carregarMotivos(),
                    carregarTiposLotePiloto(),
                    carregarStatus()
                ]);

                if (acaoFormulario === 'EDITAR' && id) {
                    document.title = 'Editar Lote Piloto #' + id;
                    $('#titulo').text('Editar Lote Piloto #' + id);

                    await carregarDadosLotePiloto();
                    $('#codigoPlanta').prop('disabled', true).trigger('change.select2');
                    $('#btnBuscarCliente, #btnPesquisarProduto').prop('disabled', true);
                    $('#grupoTipoItem').addClass('d-none');
                } else {
                    document.title = 'Cadastro de Lote Piloto';
                    $('#titulo').text('Cadastro de Lote Piloto');
                    $('#grupoLotePilotoAprovado, #grupoStatus').addClass('d-none');

                    if (acaoFormulario === 'CONTINUAR') {
                        if (!id || !['NOVO_LP', 'MESMO_PRODUTO'].includes(tipoContinuacao)) {
                            Toast.error('Parâmetros inválidos para continuar o lote piloto.');
                            $('#btnSalvar').prop('disabled', true);
                        } else {
                            await carregarDadosLotePiloto();
                            $('#indicadorContinuacao').removeClass('d-none').text('Continuação - ' + tipoContinuacao.replace('_', ' ') + ` | Cód. #${codigoProdutoAnterior}`);
                        }
                    }
                }
                
                inicializarValidacao();
                $('#codigoPlanta').trigger('change');
                atualizarJustificativas();
            } catch (error) {
                console.error('Erro ao inicializar formulário:', error);
                Toast.error('Erro ao inicializar formulário');
            } finally {
                exibirLoading(false);
            }
        });
    </script>
</head>
<body>
    <div class="container">
        <div class="d-flex flex-col mt-4 mb-2 justify-content-between align-items-center flex-wrap gap-2">
            <div class="d-flex align-items-center">
                <button type="button" onclick="navegarParaGrade()" class="btn btn-sm btn-light me-2 fs-5">
                    <i class="bi bi-arrow-left"></i>
                </button>
                <h3 id="titulo" class="mb-0"></h3>
            </div>
        </div>

        <div id="indicadorContinuacao" class="badge bg-warning text-dark mt-1 d-none"></div>

        <div id="loadingOverlay" class="position-fixed top-0 start-0 w-100 h-100 d-none justify-content-center align-items-center" style="background: rgba(255,255,255,.8); z-index: 9999;">
            <div class="spinner-border text-primary" role="status">
                <span class="visually-hidden">Carregando...</span>
            </div>
        </div>

        <div class="card p-3 mt-3">
            <form id="formularioUnificado" name="formularioUnificado" novalidate>
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
                            <option value="08">ARACATUBA</option>
                            <option value="09">BLUMENAU</option>
                            <option value="10">PACAJUS</option>
                            <option value="11">PORTO FELIZ</option>
                            <option value="12">TRES BARRAS</option>
                            <option value="13">FLORESTAL</option>
                        </select>
                    </div>
                    <div class="col-12 col-md-3">
                        <label for="codigoCliente" class="form-label">Cód. Cliente <span class="required-field me-2">*</span></label>
                        <div class="input-group">
                            <button type="button" onclick="abrirModalCliente()" class="btn btn-sm btn-secondary" id="btnBuscarCliente">
                                Buscar
                                <i class="bi bi-box-arrow-up-right"></i>
                            </button>
                            <input disabled type="text" class="form-control form-control-sm" id="codigoCliente" name="codigoCliente"/>
                        </div>
                    </div>
                    <div class="col-12 col-md-2">
                        <label for="codigoLoja" class="form-label">Loja <span class="required-field">*</span></label>
                        <input type="text" class="form-control form-control-sm" id="codigoLoja" name="codigoLoja" readonly/>
                    </div>
                    <div class="col-12 col-md-4">
                        <label for="nomeCliente" class="form-label">Nome Cliente <span class="required-field">*</span></label>
                        <input type="text" class="form-control form-control-sm" id="nomeCliente" name="nomeCliente" readonly/>
                    </div>
                    <div class="col-12 col-md-3" id="grupoTipoItem">
                        <label for="tipoItem" class="form-label">Tipo do Item <span class="required-field">*</span></label>
                        <select class="form-select form-select-sm select2" id="tipoItem" name="tipoItem"></select>
                    </div>
                    <div class="col-12 col-md-3">
                        <label for="codigoProduto" class="form-label">Produto <span class="required-field me-2">*</span></label>
                        <div class="input-group">
                            <button type="button" onclick="abrirModalProduto()" class="btn btn-sm btn-secondary" id="btnPesquisarProduto">
                                Buscar
                                <i class="bi bi-box-arrow-up-right"></i>
                            </button>
                            <input disabled type="text" class="form-control form-control-sm" id="codigoProduto" name="codigoProduto"/>
                        </div>
                    </div>
                    <div class="col-12 col-md-6">
                        <label for="referencia" class="form-label">Referência</label>
                        <input type="text" class="form-control form-control-sm" id="referencia" name="referencia" readonly/>
                    </div>
                    <div class="col-12 col-md-6">
                        <label for="quantidadePaletes" class="form-label">Qtde. de Palete <span class="required-field">*</span></label>
                        <input type="number" class="form-control form-control-sm" id="quantidadePaletes" name="quantidadePaletes"/>
                    </div>
                    <div class="col-12 col-md-6">
                        <label for="quantidadePedidos" class="form-label">Qtde. de Pedidos LP <span class="required-field">*</span></label>
                        <input value="1" type="number" class="form-control form-control-sm" id="quantidadePedidos" name="quantidadePedidos" min="1"/>
                    </div>
                    <div class="col-12 col-md-4">
                        <label for="motivo" class="form-label">Motivo <span class="required-field">*</span></label>
                        <select class="form-select form-select-sm select2" id="motivo" name="motivo"></select>
                    </div>
                    <div class="col-12 col-md-4">
                        <label for="precisaLiberacao" class="form-label">Precisa de Liberação <span class="required-field">*</span></label>
                        <select class="form-select form-select-sm select2" id="precisaLiberacao" name="precisaLiberacao">
                            <option value="" selected>Selecione...</option>
                            <option value="1">SIM</option>
                            <option value="0">NÃO</option>
                        </select> 
                    </div>
                    <div class="col-12 col-md-4">
                        <label for="tipoLP" class="form-label">Tipo de LP <span class="required-field">*</span></label>
                        <select class="form-select form-select-sm select2" id="tipoLP" name="tipoLP"></select>
                    </div>
                    <div class="col-12 col-md-4 d-none" id="grupoPedidoKiwi">
                        <label for="pedidoWRK" class="form-label">Pedido KIWI <span class="required-field">*</span></label>
                        <input type="text" class="form-control form-control-sm" id="pedidoWRK" name="pedidoWRK"/>
                    </div>
                    <div class="col-12 col-md-6" id="grupoLotePilotoAprovado">
                        <label for="lotePilotoAprovado" class="form-label">Lote Piloto Aprovado <span class="required-field">*</span></label>
                        <select class="form-select form-select-sm select2" id="lotePilotoAprovado" name="lotePilotoAprovado">
                            <option value="" selected>Selecione...</option>
                            <option value="1">SIM</option>
                            <option value="0">NÃO</option>
                        </select>
                    </div>
                        <div class="col-12 col-md-6" id="grupoStatus">
                        <label for="status" class="form-label">Status <span class="required-field">*</span></label>
                        <select class="form-select form-select-sm select2" id="status" name="status"></select>
                    </div>

                    <div id="camposJustificativa" class="row g-3"></div>

                    <div class="col-12"><hr/></div>

                    <div class="col-12 col-md-4">
                        <div class="form-check form-switch">
                            <input class="form-check-input" type="checkbox" id="acompanhamentoOnduladeira" name="acompanhamentoOnduladeira"/>
                            <label class="form-check-label" for="acompanhamentoOnduladeira">Acompanhamento Onduladeira</label>
                        </div>
                    </div>
                    <div class="col-12 col-md-4">
                        <div class="form-check form-switch">
                            <input class="form-check-input" type="checkbox" id="acompanhamentoImpressora" name="acompanhamentoImpressora"/>
                            <label class="form-check-label" for="acompanhamentoImpressora">Acompanhamento Impressoras</label>
                        </div>
                    </div>
                    <div class="col-12 col-md-4">
                        <div class="form-check form-switch">
                            <input class="form-check-input" type="checkbox" id="acompanhamentoExpedicao" name="acompanhamentoExpedicao"/>
                            <label class="form-check-label" for="acompanhamentoExpedicao">Acompanhamento Expedição</label>
                        </div>
                    </div>
                    <div class="col-12 col-md-4">
                        <div class="form-check form-switch">
                            <input class="form-check-input" type="checkbox" id="acompanhamentoCliente" name="acompanhamentoCliente"/>
                            <label class="form-check-label" for="acompanhamentoCliente">Acompanhamento Cliente</label>
                        </div>
                    </div>
                    <div class="col-12 col-md-4">
                        <div class="form-check form-switch">
                            <input class="form-check-input" type="checkbox" id="acompanhamentoDDP" name="acompanhamentoDDP"/>
                            <label class="form-check-label" for="acompanhamentoDDP">Acompanhamento DDP</label>
                        </div>
                    </div>
                    <div class="col-12 col-md-4">
                        <div class="form-check form-switch">
                            <input class="form-check-input" type="checkbox" id="acompanhamentoClicheria" name="acompanhamentoClicheria"/>
                            <label class="form-check-label" for="acompanhamentoClicheria">Acompanhamento Clicheria</label>
                        </div>
                    </div>
                </div>

                <div id="divAcoes" class="d-flex justify-content-end gap-2 mt-4 pt-3 border-top">
                    <button type="button" class="btn btn-sm btn-outline-secondary" onclick="navegarParaGrade()">
                        <i class="bi bi-arrow-left me-1"></i>
                        Voltar
                    </button>
                    <button type="button" class="btn btn-sm btn-success" id="btnSalvar">
                        <i class="bi bi-check-lg me-1"></i>
                        Salvar Lote Piloto
                    </button>
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

    <!-- Modal de Pesquisa de Produto -->
    <div class="modal fade" id="modalPesquisaProduto"  aria-labelledby="modalPesquisaProdutoLabel" aria-hidden="true">
        <div class="modal-dialog modal-lg modal-dialog-centered modal-dialog-scrollable">
            <div class="modal-content">
                <div class="modal-header">
                    <h5 class="modal-title" id="modalPesquisaProdutoLabel">
                        <i class="bi bi-search me-2"></i>Buscar Produto
                    </h5>
                    <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Fechar"></button>
                </div>
                <div class="modal-body">
                    <div class="mb-3">
                        <label for="produtoBusca" class="form-label">Pesquisar Produto</label>
                        <div class="input-group">
                            <span class="input-group-text"><i class="bi bi-search"></i></span>
                            <input type="text" 
                                   class="form-control form-control-sm" 
                                   id="produtoBusca" 
                                   placeholder="Digite o código, nome ou CNPJ para buscar..."
                                   autocomplete="off">
                        </div>
                    </div>

                    <div id="produtoResultados" style="max-height: 400px; overflow-y: auto;">
                    </div>
                </div>
                <div class="modal-footer">
                    <button type="button" class="btn btn-sm btn-outline-secondary" data-bs-dismiss="modal">
                        Cancelar
                    </button>
                    <button type="button" class="btn btn-sm btn-primary" id="btnConfirmarProduto" onclick="confirmarProduto()" disabled>
                        <i class="bi bi-check-lg me-1"></i>Selecionar
                    </button>
                </div>
            </div>
        </div>
    </div>
    
</body>
</html>
