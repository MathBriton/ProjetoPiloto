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
        #containerPedido {
        }
    </style>

    <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.0.2/dist/js/bootstrap.bundle.min.js" integrity="sha384-MrcW6ZMFYlzcLA8Nl+NtUVF0sA7MsXsP1UyJoMp4YLEuNSfAP+JcXn/tWtIaxVXM" crossorigin="anonymous"></script>
    <script src="https://cdn.jsdelivr.net/npm/select2@4.1.0-rc.0/dist/js/select2.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/bootstrap-datepicker@1.10.0/dist/js/bootstrap-datepicker.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/bootstrap-datepicker@1.10.0/dist/locales/bootstrap-datepicker.pt-BR.min.js"></script>

    <script src="/corporativo/components/toast.js?v=<%=Now()%>"></script>
    <script src="/corporativo/components/confirmation.js?v=<%=Now()%>"></script>
    <script src="/corporativo/components/upload/fileUpload.js?v=<%=Now()%>"></script>

    <script>
        let pedido = {}

        // Listas usadas para preencher campos dinâmicos
        let listaTiposFrete = []
        let listaStatus = []
        let listaCondicoesPagamento = []
        let listaFiliais = []

        function carregarTiposAlteracao() {
            const $select = $('#tipoAlteracao');
            $select.empty().append('<option value="">Selecione o tipo de alteração...</option>');

            const params = {
                acao: 'listarTiposAlteracao'
            };

            $.ajax({
                url: 'SolicitacaoAlteracaoPedidoRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: params,
                dataType: 'json',
                success: function(response) {
                    if (!response.sucesso) {
                        Toast.error("Erro ao carregar tipos de alteração: " + response.mensagem);
                        return;
                    }
                    
                    response.dados.forEach(function(item) {
                        $select.append($('<option>', { value: item.id, text: item.descricao }));
                    });
                },
                error: function(xhr, status, error) {
                    Toast.error('Erro ao carregar tipos de alteração:', error);
                }
            });
        }

        function carregarTiposFrete() {
            const params = {
                acao: 'listarTiposFrete'
            };

            $.ajax({
                url: 'SolicitacaoAlteracaoPedidoRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: params,
                dataType: 'json',
                success: function(response) {
                    if (!response.sucesso) {
                        Toast.error("Erro ao carregar setores: " + response.mensagem);
                        return;
                    }

                    listaTiposFrete = response.dados
                },
                error: function(xhr, status, error) {
                    Toast.error('Erro ao carregar setores:', error);
                }
            });
        }

        function carregarStatus() {
            const params = {
                acao: 'listarStatus'
            };

            $.ajax({
                url: 'SolicitacaoAlteracaoPedidoRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: params,
                dataType: 'json',
                success: function(response) {
                    if (!response.sucesso) {
                        Toast.error("Erro ao carregar setores: " + response.mensagem);
                        return;
                    }

                    listaStatus = response.dados
                },
                error: function(xhr, status, error) {
                    Toast.error('Erro ao carregar setores:', error);
                }
            });
        }

        function carregarFiliais() {
            const params = {
                acao: 'listarFiliais'
            };

            $.ajax({
                url: 'SolicitacaoAlteracaoPedidoRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: params,
                dataType: 'json',
                success: function(response) {
                    if (!response.sucesso) {
                        Toast.error("Erro ao carregar filiais: " + response.mensagem);
                        return;
                    }

                    listaFiliais = response.dados
                },
                error: function(xhr, status, error) {
                    Toast.error('Erro ao carregar filiais:', error);
                }
            });
        }

        function carregarCondicoesPagamento() {
            const params = {
                acao: 'listarCondicoesPagamento'
            };

            $.ajax({
                url: 'SolicitacaoAlteracaoPedidoRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: params,
                dataType: 'json',
                success: function(response) {
                    if (!response.sucesso) {
                        Toast.error("Erro ao carregar condições de pagamento: " + response.mensagem);
                        return;
                    }

                    listaCondicoesPagamento = response.dados
                },
                error: function(xhr, status, error) {
                    Toast.error('Erro ao carregar condições de pagamento:', error);
                }
            });
        }

        function pesquisarPedido() {
            const baseUrl = `${location.protocol}//${location.host}`;
            const url = `BuscaPedido2.asp?IdUsuario=1`;
            window.open(url, '', "width=650,height=400");
        }
        

        function obterPedido() {
            const params = {
                acao: 'obterPedido',
                codPedido: $('#txtPedido').val()?.trim()
            };

            resetarCamposPedido()
            exibirLoadingPedido()

            $.ajax({
                url: 'SolicitacaoAlteracaoPedidoRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: params,
                dataType: 'json',
                success: function(response) {
                    if (!response.sucesso) {
                        exibirPedidoNaoEncontrado();
                        return;
                    }

                    pedido = {
                        ...response.dados, 
                        permiteEntregaFinalDeSemana: response.dados.permiteEntregaFinalDeSemana === '0' ? false : true
                    }

                    exibirPedidoEncontrado();
                },
                error: function(xhr, status, error) {
                    exibirPedidoNaoEncontrado();
                }
            });
        }

        function resetarCamposPedido() {
            pedido = {}

            // Limpa e desabilita Campos
            $('#tipoAlteracao').prop('disabled', true).val('').trigger('change');
            $('#observacao').prop('disabled', true).val('')
            $('#containerCamposAlteracao').html('');

            const uploadEvidencia = $('#uploadEvidencia')[0];
            uploadEvidencia.setDisabled(true)
        }

        function exibirLoadingPedido() {
            let html = `
               <div class="text-center">
                    <div class="spinner-border text-primary text-center" role="status">
                        <span class="visually-hidden">Loading...</span>
                    </div>
               </div>
            `
            $('#containerPedido').html(html);
        }

        function exibirPedidoEncontrado() {
            $('#containerPedido').html(`
                <div class="row g-3">
                    <div class="col-md-6">
                        <label for="nomeCliente" class="form-label">Nome cliente</label>
                        <input disabled value="${pedido.nomeCliente}" class="form-control form-control-sm" name="nomeCliente" id="nomeCliente" />
                    </div>
                    <div class="col-md-3">
                        <label for="codCliente" class="form-label">Cod. Cliente</label>
                        <input disabled value="${pedido.codCliente}" class="form-control form-control-sm" name="codCliente" id="codCliente" />
                    </div>
                    <div class="col-md-3">
                        <label for="codLoja" class="form-label">Loja Cliente</label>
                        <input disabled value="${pedido.codLoja}" class="form-control form-control-sm" name="codLoja" id="codLoja" />
                    </div>
                    <div class="col-md-6">
                        <label for="referencia" class="form-label">Referência</label>
                        <input disabled value="${pedido.referencia}" type="text" class="form-control form-control-sm" name="referencia" id="referenciaProduto" />
                    </div>
                    <div class="col-md-3">
                        <label for="codPedido" class="form-label">Cod. Pedido</label>
                        <input disabled value="${pedido.codPedido}" type="text" class="form-control form-control-sm" name="codPedido" id="codPedido" />
                    </div>
                    <div class="col-md-3">
                        <label for="codProduto" class="form-label">Cod. Produto</label>
                        <input disabled  value="${pedido.codProduto}" type="text" class="form-control form-control-sm" name="codProduto" id="codProduto" />
                    </div>
                </div>

                <div class="border-bottom border-1 pb-4"/>
            `);

            // Libera campos Bloqueados
            $('#tipoAlteracao').prop('disabled', false);
            $('#observacao').prop('disabled', false);

            const uploadEvidencia = $('#uploadEvidencia')[0];
            uploadEvidencia.setDisabled(false)
        }

        function exibirPedidoNaoEncontrado() {
            $('#containerPedido').html('<div class="alert alert-danger text-center"> <i class="bi bi-exclamation-triangle me-1"></i> Pedido não encontrado </div>');
        }
        
        function anexarEventos() {
            $('#formularioSolicitacao').on('submit', function(e) {
                e.preventDefault();
                salvarFormulario();
            });

            $('#tipoAlteracao').on('change', campoTipoAlteracaoAlterado);
            $('#uploadEvidencia').on('error', (e) => { Toast.error(e.detail.message); });
        }

        function campoTipoAlteracaoAlterado() {
            $('#containerCamposAlteracao').html('');

            const valorTipoAlteracao = $('#tipoAlteracao').val();
            const codPedido = $('#codPedido').val();
            
            if (!valorTipoAlteracao) return;
            if (!codPedido) return;
            
            renderizarCamposAlteracao(valorTipoAlteracao);
            inicializarSelect2();
        }

        function renderizarCamposAlteracao(tipoAlteracao) {
            switch(tipoAlteracao) {
                case '1':
                    // Frete
                    renderizarCamposSelect(listaTiposFrete, pedido.frete);
                    break;
                case '2':
                    // Preço
                    renderizarCamposNumerico(pedido.preco);
                    break;
                case '3':
                    // Quantidade
                    renderizarCamposNumerico(pedido.quantidade, true);
                    break;
                case '4':
                    // QuantidadePronto
                    renderizarCamposNumerico(pedido.quantidadePronto);
                    break;
                case '5':
                    // Data de Entrega
                    renderizarCamposData(pedido.dataEntregaDe, pedido.dataEntregaAte);
                    break;
                case '6':
                    // Condição de Pagamento
                    renderizarCamposSelect(listaCondicoesPagamento, pedido.codCondicaoPagamento);
                    break;
                case '7':
                    // Cancelar
                    renderizarCamposSelect(listaStatus, pedido.status, 'C');
                    break;
                case '8':
                    // Pedido do Cliente
                    renderizarCamposTexto(pedido.pedidoCliente);
                    break;
                case '9':
                    // Status
                    renderizarCamposSelect(listaStatus, pedido.status);
                    break;
                case '10':
                    // Filial/Fabrica
                    renderizarCamposSelect(listaFiliais, pedido.codFilial);
                    break;
                case '14':
                    // Pedido do Cliente
                    renderizarCamposNumerico(pedido.pedidoCliente);
                    break;
                default:
                    break;
            }
        }

       function renderizarCamposSelect(listaItens, valorSelecionadoDe = null, valorSelecionadoPara = null) {
            const listaDe = listaItens;
            const listaPara = listaItens.filter(item => item.id != valorSelecionadoDe);

            const gerarOptions = (lista) => 
                lista.map(item => `<option value="${item.id}">${item.descricao}</option>`).join('');

            $('#containerCamposAlteracao').html(`
                <div class="row g-3">
                    <div class="col-md-6">
                        <label for="campoDe" class="form-label">De</label>
                        <select ${valorSelecionadoDe ? 'disabled' : ''} 
                                class="form-select form-select-sm select2" 
                                id="campoDe" name="campoDe" required>
                            <option value="">Selecione o valor atual...</option>
                            ${gerarOptions(listaDe)}
                        </select>
                    </div>
                    <div class="col-md-6">
                        <label for="campoPara" class="form-label">Para</label>
                        <select ${valorSelecionadoPara ? 'disabled' : ''} 
                                class="form-select form-select-sm select2" 
                                id="campoPara" name="campoPara" required>
                            <option value="">Selecione o novo valor...</option>
                            ${gerarOptions(listaPara)}
                        </select>
                    </div>
                </div>
            `);

            if (valorSelecionadoDe) $('#campoDe').val(valorSelecionadoDe).trigger('change');
            if (valorSelecionadoPara) $('#campoPara').val(valorSelecionadoPara).trigger('change');
        }

        function renderizarCamposTexto(valorSelecionado = null) { 
            $('#containerCamposAlteracao').html(`
                <div class="row g-3">
                    <div class="col-md-6">
                        <label for="campoDe" class="form-label">De</label>
                        <input disabled type="text" class="form-control form-control-sm" 
                            id="campoDe" name="campoDe" required>
                    </div>
                    <div class="col-md-6">
                        <label for="campoPara" class="form-label">Para</label>
                        <input type="text" class="form-control form-control-sm" 
                            id="campoPara" name="campoPara" required>
                    </div>
                </div>
            `);
        }

        function renderizarCamposNumerico(valorSelecionado = null, campoCalculado = false) {
            const layoutPadrao = `
                <div class="row g-3">
                    <div class="col-md-6">
                        <label for="campoDe" class="form-label">De</label>
                        <input disabled type="number" class="form-control form-control-sm" 
                            id="campoDe" name="campoDe" step="0.00001" min="0" max="100" required>
                    </div>
                    <div class="col-md-6">
                        <label for="campoPara" class="form-label">Para</label>
                        <input type="number" class="form-control form-control-sm" 
                            id="campoPara" name="campoPara" step="0.00001" min="0" required>
                    </div>
                </div>
            `;

            const layoutCalculado = `
                <div class="row g-3">
                    <div class="col-md-4">
                        <label for="campoDe" class="form-label">De</label>
                        <input disabled type="number" class="form-control form-control-sm" 
                            id="campoDe" name="campoDe" step="0.00001" min="0" required>
                    </div>
                    <div class="col-md-4">
                        <label for="campoPara2" class="form-label">Quantidade</label>
                        <input type="number" class="form-control form-control-sm" 
                            id="campoPara2" name="campoPara2" step="0.00001" min="0" required>
                    </div>
                    <div class="col-md-4">
                        <label for="campoPara" class="form-label">Para</label>
                        <input disabled type="number" class="form-control form-control-sm" 
                            id="campoPara" name="campoPara" step="0.00001" min="0" required>
                    </div>
                </div>
            `;

            $('#containerCamposAlteracao').html(campoCalculado ? layoutCalculado : layoutPadrao);

            if (valorSelecionado) {
                $('#campoDe').val(valorSelecionado.replace(',', '.'));
            }

            if (campoCalculado) {
                $('#campoPara2').on('input', function() {
                    const quantidade = parseFloat($(this).val()) || 0;
                    const valorCalculado = quantidade * ((pedido.tolerancia / 100) + 1); 
                    $('#campoPara').val(parseInt(valorCalculado));
                });
            }
        }

        function renderizarCamposData(valorSelecionadoDe = null, valorSelecionadoAte = null) {
            $('#containerCamposAlteracao').html(`
                <div class="row g-3">
                    <div class="col-md-4">
                        <label for="campoDe" class="form-label">Data atual</label>
                        <input disabled type="text" class="form-control form-control-sm" 
                            id="campoDe" name="campoDe" required>
                    </div>
                    <div class="col-md-4">
                        <label for="campoPara" class="form-label">Nova Data (De)</label>
                        <input type="date" class="form-control form-control-sm" 
                            id="campoPara" name="campoPara" placeholder="dd/mm/aaaa" required>
                    </div>
                    <div class="col-md-4">
                        <label for="campoPara2" disabled class="form-label">Nova Data (Até)</label>
                        <input type="date" class="form-control form-control-sm" 
                            id="campoPara2" name="campoPara2" placeholder="dd/mm/aaaa" disabled required>
                    </div>
                </div>
            `);

            if (valorSelecionadoDe && valorSelecionadoAte) {
                $('#campoDe').val(`${formatarDataBR(valorSelecionadoDe)} até ${formatarDataBR(valorSelecionadoAte)}`);
            }

            $('#campoPara').on('change', function(e) {
                const dataInicio = new Date($(this).val() + ' 00:00:00');
                const dataFinal = adicionarDiasUteis(dataInicio, pedido.intervaloDataEntrega, pedido.permiteEntregaFinalDeSemana);    
                $('#campoPara2').val(dataFinal);
            });
        }

        function adicionarDiasUteis(data, quantidadeDias, permiteFinalDeSemana = false) {
            const novaData = data;
            quantidadeDias = parseInt(quantidadeDias);
            
            if (permiteFinalDeSemana) {
                novaData.setDate(novaData.getDate() + quantidadeDias);
            } else {
                let diasAdicionados = 0;
                
                while (diasAdicionados < quantidadeDias) {
                    novaData.setDate(novaData.getDate() + 1);
                    
                    const diaSemana = novaData.getDay();
                    const ehDiaUtil = diaSemana !== 0 && diaSemana !== 6;
                    
                    if (ehDiaUtil) {
                        diasAdicionados++;
                    }
                }
            }
            
            const anoFinal = novaData.getFullYear();
            const mesFinal = (novaData.getMonth() + 1).toString().padStart(2, '0');
            const diaFinal = novaData.getDate().toString().padStart(2, '0');
            
            return `${anoFinal}-${mesFinal}-${diaFinal}`;
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

        function navegarParaGrade() {
            window.location.href = 'SolicitaAlteracao.asp'
        }

        function salvarFormulario() {
            const uploadEvidencia = document.getElementById('uploadEvidencia');
            const evidencias = uploadEvidencia.getFiles();

            const $btnSalvar = $('#btnSalvar');
            const $loadingIcon = $btnSalvar.find('#loadingIcon');
            const $checkIcon = $btnSalvar.find('#checkIcon');

            $btnSalvar.prop('disabled', true);
            $checkIcon.css('display', 'none');
            $loadingIcon.css('display', 'inline-block');

            const valorPara = $('#tipoAlteracao').val() == 5 ? 
                formatarDataUS($('#campoPara').val()) : $('#campoPara').val();

            const valorPara2 = $('#tipoAlteracao').val() == 5 ? 
                formatarDataUS($('#campoPara2').val()) : $('#campoPara2').val();

            const dadosFormulario = {
                acao: 'salvarSolicitacaoAlteracao',
                codFilial: pedido.codFilial,
                codPedido: $('#codPedido').val(),
                tipoAlteracao: $('#tipoAlteracao').val(),
                valorDe : $('#campoDe').val(),
                valorPara: valorPara,
                valorPara2: valorPara2,
                observacao: $('#observacao').val(),
                evidencias: JSON.stringify(evidencias)
            };

            $.ajax({
                url: 'SolicitacaoAlteracaoPedidoRepositorio.asp',
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

        function formatarDataBR(str) {
            return str.split('-').reverse().join('/');
        }

        function formatarDataUS(str) {
            return str.split('/').reverse().join('-');
        }

        $(document).ready(function() {
            inicializarSelect2();
            carregarTiposAlteracao();
            carregarTiposFrete();
            carregarCondicoesPagamento();
            carregarStatus();
            carregarFiliais();
            anexarEventos();
        });
    </script>
</head>
<body>
    <div class="container my-4">
        <div class="d-flex my-4">
            <button onclick="navegarParaGrade()" class="btn btn-sm btn-light me-2 fs-5"> 
                <i class="bi bi-arrow-left"></i> 
            </button>
            <h3>Solicitar Alteração de Pedido</h3>
        </div>

        <div class="card p-3">
            <form id="formularioSolicitacao" name="formularioSolicitacao">
                <div class="row g-3">
                    <div class="col-12">
                        <label for="txtPedido" class="form-label">Pedido <span class="required-field">*</span></label>
                        <div class="input-group input-group-sm">
                            <button type="button" class="btn btn-sm btn-secondary" onclick="pesquisarPedido()">
                                Buscar <i class="bi bi-box-arrow-up-right"></i>
                            </button>
                            <input type="text" class="form-control form-control-sm" id="txtPedido" name="txtPedido" required>
                            <button type="button" class="btn btn-sm btn-secondary" onclick="obterPedido()">
                                <i class="bi bi-search"></i>
                            </button>
                        </div>
                        <small class="text-muted">Clique no botão para buscar ou digite o código</small>
                    </div>

                    <div id="containerPedido">
                        <!-- Mostra resultado da pesquisa de pedido -->
                    </div>

                    <div class="col-12">
                        <label for="tipoAlteracao" class="form-label">Tipo de Alteração <span class="required-field">*</span></label>
                        <select required disabled class="form-select form-select-sm select2" id="tipoAlteracao" name="tipoAlteracao" required>
                            <option value="">Selecione o tipo...</option>
                        </select>
                    </div>

                    <div id="containerCamposAlteracao">
                        <!-- Campos gerados dinamicamente -->
                    </div>

                    <div class="col-12">
                        <label for="observacao" class="form-label">Observação <span class="required-field">*</span></label>
                        <textarea required disabled required class="form-control form-control-sm" id="observacao" name="observacao" rows="4"></textarea>
                    </div>

                    <div class="col-12 mt-4">
                        <label class="form-label">Evidência</label>
                        <file-upload 
                            id="uploadEvidencia"
                            file-code="ALTERACAO_PEDIDO"
                            max-files="5"
                            disabled
                            required
                            multiple>
                        </file-upload>
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
