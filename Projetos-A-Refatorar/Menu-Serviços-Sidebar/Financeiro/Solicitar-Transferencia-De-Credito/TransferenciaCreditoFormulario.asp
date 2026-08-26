<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file = "../../cnn2/cnn2.asp" -->
<!--#include File = "../../funcoes/funcoes.asp"-->
<!--#include file="../../includes/variaveis.asp" -->
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
    <title>Smurfit Kappa - Transferência de Limite de Crédito</title>
    <meta charset="utf-8"/>
    <meta name="viewport" content="width=device-width, initial-scale=1"/>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.0.2/dist/css/bootstrap.min.css" rel="stylesheet" integrity="sha384-EVSTQN3/azprG1Anm3QDgpJLIm9Nao0Yz1ztcQTwFspd3yD65VohhpuuCOmLASjC" crossorigin="anonymous" />
    <link rel="stylesheet" href="https://use.fontawesome.com/releases/v5.7.0/css/all.css" integrity="sha384-lZN37f5QGtY3VHgisS14W3ExzMWZxybE1SJSEsQp9S+oqd12jhcu+A56Ebc1zFSJ" crossorigin="anonymous" />
    <link href="https://cdn.jsdelivr.net/npm/select2@4.1.0-rc.0/dist/css/select2.min.css" rel="stylesheet" />
    <link href="https://cdn.jsdelivr.net/npm/select2-bootstrap-5-theme@1.3.0/dist/select2-bootstrap-5-theme.min.css" rel="stylesheet" />
        <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.1/font/bootstrap-icons.css" />

    <style>
        body {
            background-color: #f5f5f5;
            font-size: .9rem;
        }
        .btn-manual {
            display: flex;
            justify-content: center;
            align-items: center;
            color: #fff;
            background: #0d6efd;
            border: 1px solid #0d6efd;
            width: 35px;
            height: 35px;
            border-radius: 7px;
            margin-left: 12px;
            font-size: 1.2rem;
        }
        .card {
            padding: 1rem;
        }
        .divider {
            height: 1px;
            background-color: #dee2e6;
            margin: 20px 0;
        }
        .form-label {
            font-weight: 600;
        }
        .required-field {
            color: #dc3545;
        }
        .mensagem-erro, .mensagem-sucesso {
            padding: 20px;
            border-radius: 8px;
            text-align: center;
        }
        .mensagem-erro {
            background-color: #f8f5d7ff;
            border: 1px solid #dcb535ff;
            color: #846420ff;
        }
        .mensagem-erro i {
            font-size: 2rem;
            margin-bottom: 15px;
            color: #a07a29ff;
        }
        .mensagem-sucesso {
            background-color: #d1e7dd;
            border: 1px solid #198754;
            color: #0f5132;
            margin: 40px auto;
            max-width: 600px;
        }
        .mensagem-sucesso i {
            font-size: 3rem;
            margin-bottom: 15px;
        }
        .mensagem-sucesso h4 {
            margin-bottom: 10px;
        }
        .loading-container {
            text-align: center;
            padding: 30px;
        }
        .loading-container .spinner-border {
            width: 3rem;
            height: 3rem;
            color: #0d6efd;
        }
        .loading-container p {
            margin-top: 15px;
            color: #6c757d;
        }
        .transfer-container {
            display: flex;
            gap: 2rem;
            align-items: stretch;
        }
        .transfer-card {
            flex: 1;
            background: #fafafa;
            border: 1px solid #e0e0e0;
            border-radius: 8px;
            padding: .7rem;
        }
        .transfer-header {
            display: flex;
            justify-content: space-between;
            align-items: baseline;
            margin-bottom: 1rem;
            padding-bottom: 0.75rem;
            border-bottom: 1px solid #e0e0e0;
        }
        .transfer-label {
            font-size: 0.75rem;
            color: #757575;
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }
        .transfer-title {
            font-size: 1rem;
            font-weight: 600;
            color: #212121;
        }
        .transfer-details {
            display: flex;
            flex-direction: column;
            gap: 0.4rem;
        }
        .transfer-row {
            display: flex;
            justify-content: space-between;
            align-items: center;
        }
        .transfer-row-label {
            font-size: 0.875rem;
            color: #616161;
        }
        .transfer-row-value {
            font-size: 0.875rem;
            font-weight: 600;
            color: #212121;
        }
        .transfer-row.highlight {
            padding-top: 0.75rem;
            border-top: 1px solid #e0e0e0;
        }
        .transfer-row.highlight .transfer-row-value {
            font-size: 1rem;
            color: #198754;
        }
        .transfer-arrow-container {
            display: flex;
            flex-direction: column;
            align-items: center;
            justify-content: center;
            gap: 0.5rem;
            min-width: 80px;
        }
        .transfer-arrow {
            width: 32px;
            height: 32px;
            border: 1px solid #bdbdbd;
            border-radius: 50%;
            display: flex;
            align-items: center;
            justify-content: center;
            background: white;
        }
        .transfer-arrow i {
            color: #757575;
            font-size: 0.875rem;
        }
        @media (max-width: 991px) {
            .transfer-container {
                flex-direction: column;
            }
            .transfer-arrow-container {
                flex-direction: row;
                min-width: auto;
                padding: 1rem 0;
            }
            .transfer-arrow {
                transform: rotate(90deg);
            }
        }
        
        .btn-selecionar {
            display: flex;
            justify-content: center;
            align-items: baseline;
            padding: .25rem .5rem;
            font-size: .875rem;
            border: 1px solid #dee2e6;
            background-color: #fff;
            color: #6c757d;
            border-radius: 4px;
            cursor: pointer;
        }

        .btn-selecionar:hover {
            background-color: #0d6efd;
            color: #fff;
            border-color: #0d6efd;
        }

        .btn-selecionar:disabled {
            opacity: 0.5;
            cursor: not-allowed;
        }

        .btn-selecionar.selecionado {
            background-color: #198754;
            color: #fff;
            border-color: #198754;
            cursor: default;
        }

        .btn-selecionar.selecionado:hover {
            background-color: #198754;
            border-color: #198754;
        }
        .cliente-code {
            font-weight: 600;
            color: #212529;
        }
        .cliente-detail {
            font-size: 0.875rem;
            color: #6c757d;
            margin-bottom: 0.25rem;
        }
        .cliente-detail strong {
            color: #495057;
        }
    </style>

    <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.0.2/dist/js/bootstrap.bundle.min.js" integrity="sha384-MrcW6ZMFYlzcLA8Nl+NtUVF0sA7MsXsP1UyJoMp4YLEuNSfAP+JcXn/tWtIaxVXM" crossorigin="anonymous"></script>
    <script src="https://cdn.jsdelivr.net/npm/select2@4.1.0-rc.0/dist/js/select2.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/jquery-maskmoney@3.0.2/dist/jquery.maskMoney.min.js"></script>        
    <script src="../util/Util.js?v=<%=Now()%>"></script>
    <script src="/corporativo/funcoes/funcoesJAVASCRIPT.js?v=<%=Now()%>"></script>

    <script>
        let preview = { plantaOrigem: null, plantaDestino: null, limiteOrigem: 0, limiteDestino: 0, valor: 0, saldoOrigem: 0 };
        let clientesDestino = [];
        let clienteSelecionado = null;

        const PLANTAS = {
            '06': 'Maranguape', 
            '00': 'Pirapetinga', 
            '04': 'Uberaba', 
            '05': 'Bento', 
            '07': 'Saquarema'
        };
       
        function PesquisaCliente() {
            const codPlanta = $('#plantaOrigem').val();
            if (!codPlanta) return alert('Selecione a planta primeiro');
            const url = `${location.protocol}//${location.host}/corporativo/buscaCLIENTE.asp?id1=codClienteOrigem&id2=codLojaOrigem&id3=nomeCliente&id6=${codPlanta}`;
            window.open(url, '', "width=650,height=400");
        }

        function plantaAlterada() {
            limparDadosCliente();
            $('#divDadosCliente').hide();
        }

        function obterDadosTransferencia() {
            const dados = {
                plantaOrigem: $('#plantaOrigem').val(),
                plantaDestino: $('#plantaDestino').val(),
                codClienteOrigem: $('#codClienteOrigem').val(),
                codLojaOrigem: $('#codLojaOrigem').val() ?? "01"
            };

            if (!dados.plantaOrigem || !dados.plantaDestino || !dados.codClienteOrigem) return;

            $('#divLoadingCliente').show();
            $('#divDadosCliente').hide();
            $('#divErroCliente').hide();

            $.ajax({
                url: 'TransferenciaCreditoRepositorio.asp?v=' + Math.random(),
                type: 'POST',
                data: { acao: "obterDadosTransferencia", ...dados },
                dataType: 'json',
                success: function(response) {
                    if (response.sucesso) {
                        processarDadosCliente(response.dados, dados);
                    } else {
                        mostrarErro(response.mensagem);
                    }
                },
                error: function() {
                    mostrarErro('Erro ao carregar dados do cliente. Por favor, tente novamente.');
                }
            });
        }

        function processarDadosCliente(dados, params) {
            preview.lojaOrigem = params.codLojaOrigem;
            preview.plantaOrigem = params.plantaOrigem;
            preview.plantaDestino = params.plantaDestino;
            preview.limiteOrigem = dados.limiteAtual || 0;
            preview.saldoOrigem = dados.saldo || 0;
            clientesDestino = dados.clientesDestino || [];

            $('#nomeCliente').val(dados.nomeCliente.trim());
            $('#limiteAtualOrigem').text(formatarMoeda(dados.limiteAtual));
            $('#limiteAbertoOrigem').text(formatarMoeda(dados.aberto));
            $('#saldoOrigem').text(formatarMoeda(dados.saldo));
            $('#previewOriginName').text(PLANTAS[params.plantaOrigem]);
            $('#previewOriginCode').text(params.codClienteOrigem);
            $('#previewOriginLoja').text(params.codLojaOrigem || '-');

            if (clientesDestino.length === 0) {
                mostrarErro('Nenhum cliente de destino disponível para transferência');
            } else if (clientesDestino.length === 1) {
                clienteSelecionado = clientesDestino[0];
                preencherDadosDestino(clienteSelecionado);
                preview.limiteDestino = clienteSelecionado.limiteAtual;
                $('#divDadosCliente').show();
                $('#divLoadingCliente').hide();
                $('#divErroCliente').hide();
                atualizarPreview();
            } else {
                $('#divDadosCliente').show();
                $('#divLoadingCliente').hide();
                $('#divErroCliente').hide();
                abrirModalSelecaoDestino();
            }
        }

        function preencherDadosDestino(cliente) {
            $('#limiteAtualDestino').text(formatarMoeda(cliente.limiteAtual || 0));
            $('#limiteAbertoDestino').text(formatarMoeda(cliente.aberto || 0));
            $('#saldoDestino').text(formatarMoeda(cliente.saldo || 0));
            $('#previewDestName').text(PLANTAS[preview.plantaDestino]);
            $('#previewDestCode').text((cliente.codigo || '').trim());
            $('#previewDestLoja').text((cliente.loja || '').trim());
            
            $('#hiddenCodClienteDestino').val((cliente.codigo || '').trim());
            $('#hiddenCodLojaDestino').val((cliente.loja || '').trim());
        }


        function abrirModalSelecaoDestino () {
            new bootstrap.Modal(document.getElementById('modalSelecaoDestino')).show();
            const $container = $('#listaClientesDestino');

            let html = `
                <div class="table-responsive">
                    <table class="table table-sm table-hover mb-0">
                        <thead class="thead-light">
                            <tr>
                                <th style="width: 50px;"></th>
                                <th>Nome</th>
                                <th>Código</th>
                                <th>Loja</th>
                                <th>Estado</th>
                                <th>CNPJ</th>
                                <th>Limite Atual</th>
                                <th>Em Aberto</th>
                                <th>Saldo</th>
                            </tr>
                        </thead>
                        <tbody>
            `;

            clientesDestino.forEach((cliente, index) => {
                html += `
                    <tr data-index="${index}">
                        <td class="text-center">
                            <button onclick="selecionarClienteDestino(${index})"
                                class="btn-selecionar" 
                                title="Selecionar este item">
                                <i class="bi bi-circle"></i> 
                            </button>
                        </td>
                        <td>${cliente.nome}</td>
                        <td>${cliente.codigo}</td>
                        <td>${cliente.loja}</td>
                        <td>${cliente.estado}</td>
                        <td>${formatarCNPJ(cliente.cnpj)}</td>
                        <td>${cliente.limiteAtual}</td>
                        <td>${cliente.aberto}</td>
                        <td>${cliente.saldo}</td>
                    </tr>
                `;
            });

            html += `
                        </tbody>
                    </table>
                </div>
            `;
            
            $container.html(html);
        }

        function selecionarClienteDestino(index) {
            clienteSelecionado = clientesDestino[index];
            
            $('#listaClientesDestino .btn-selecionar')
                .removeClass('selecionado')
                .html('<i class="bi bi-circle"></i>')
                .attr('title', 'Selecionar este item');
            
            $('#listaClientesDestino tr[data-index="' + index + '"] .btn-selecionar')
                .addClass('selecionado')
                .html('<i class="bi bi-check-circle-fill"></i>')
                .attr('title', 'Item Selecionado');
            
            $('#btnConfirmarDestino').prop('disabled', false);
        }

        /* function abrirModalSelecaoDestino() {
            const container = $('#listaClientesDestino').empty();
            
            clientesDestino.forEach((cliente, index) => {
                container.append(`
                    <div class="cliente-item" data-index="${index}">
                        <div class="cliente-item-header">
                            <span class="cliente-code">Código: ${(cliente.codigo || '').trim()} - Loja: ${(cliente.loja || '').trim()}</span>
                        </div>
                        <div class="cliente-detail"><strong>Limite Atual:</strong> ${formatarMoeda(cliente.limiteAtual || 0)}</div>
                        <div class="cliente-detail"><strong>Em Aberto:</strong> ${formatarMoeda(cliente.aberto || 0)}</div>
                        <div class="cliente-detail"><strong>Saldo:</strong> ${formatarMoeda(cliente.saldo || 0)}</div>
                    </div>
                `);
            });

            $('.cliente-item').on('click', function() {
                $('.cliente-item').removeClass('selected');
                $(this).addClass('selected');
                clienteSelecionado = clientesDestino[$(this).data('index')];
                $('#btnConfirmarDestino').prop('disabled', false);
            });

            new bootstrap.Modal(document.getElementById('modalSelecaoDestino')).show();
        } */

        function confirmarSelecaoDestino() {
            if (clienteSelecionado) {
                preencherDadosDestino(clienteSelecionado);
                preview.limiteDestino = clienteSelecionado.limiteAtual;
                atualizarPreview();
                bootstrap.Modal.getInstance(document.getElementById('modalSelecaoDestino')).hide();
            }
        }

        function atualizarPreview() {
            const valor = parseFloat(desformatarMoeda($('#valorTransferencia').val() || '0'));
            preview.valor = valor;

            $('#previewOriginAfter').text(formatarMoeda(preview.limiteOrigem - valor));
            $('#previewDestAfter').text(formatarMoeda(preview.limiteDestino + valor));
        }

        function salvarFormulario() {
            const $btnSalvar = $('#btnSalvar');
            $btnSalvar.prop('disabled', true);
            $('#iconCheck').hide();
            $('#iconLoading').show();

            $.ajax({
                url: 'TransferenciaCreditoRepositorio.asp',
                type: 'POST',
                data: {
                    acao: "criarTransferencia",
                    plantaOrigem: $('#plantaOrigem').val(),
                    codClienteOrigem: $('#codClienteOrigem').val(),
                    plantaDestino: $('#plantaDestino').val(),
                    codClienteDestino: $('#hiddenCodClienteDestino').val(),
                    codLojaDestino: $('#hiddenCodLojaDestino').val() ?? "01",
                    codLojaOrigem: $('#codLojaOrigem').val() ?? "01",
                    valorTransferencia: desformatarMoeda($('#valorTransferencia').val())
                },
                dataType: 'json',
                success: function(response) {
                    if (response.sucesso) {
                        $('#mensagemSucesso').show();
                        $('#containerFormulario').hide();
                    } else {
                        alert('Erro: ' + response.mensagem);
                        resetarBotaoSalvar();
                    }
                },
                error: function(xhr, status, error) {
                    alert('Erro ao processar a requisição: ' + error);
                    resetarBotaoSalvar();
                }
            });
        }

        function resetarBotaoSalvar() {
            $('#btnSalvar').prop('disabled', false);
            $('#iconLoading').hide();
            $('#iconCheck').show();
        }

        function limparDadosCliente() {
            $('#codClienteOrigem').val('');
            $('#codLojaOrigem').val('01');
            $('#nomeCliente').val('');
            $('#limiteAtualOrigem').text('');
            $('#limiteAbertoOrigem').text('');
            $('#saldoOrigem').text('');
            $('#limiteAtualDestino').text('');
            $('#limiteAbertoDestino').text('');
            $('#saldoDestino').text('');
            $('#valorTransferencia').val('');
            $('#hiddenCodClienteDestino').val('');
            $('#hiddenCodLojaDestino').val('');
            $('#divDadosCliente').hide();
            $('#divLoadingCliente').hide();
            $('#divErroCliente').hide();
            
            preview.limiteOrigem = 0;
            preview.limiteDestino = 0;
            preview.valor = 0;
            preview.saldoOrigem = 0;
            preview.plantaOrigem = null;
            preview.plantaDestino = null;
            clientesDestino = [];
            clienteSelecionado = null;
        }

        function mostrarErro(mensagem) {
            $('#mensagemErroCliente').text(mensagem);
            $('#divErroCliente').show();
            $('#divLoadingCliente').hide();
            $('#divDadosCliente').hide();
        }

        $(document).ready(function() {
            $('#plantaOrigem, #plantaDestino').select2({
                theme: 'bootstrap-5',
                placeholder: 'Selecione uma planta',
                allowClear: true,
                width: '100%'
            });

            $("#valorTransferencia").maskMoney({ thousands: '.', decimal: ',' });
            $('#valorTransferencia').on('keyup change', atualizarPreview);

            $('#formularioTransferencia').on('submit', function(e) {
                e.preventDefault();

                if (!this.checkValidity()) {
                    e.stopPropagation();
                    $(this).addClass('was-validated');
                    return false;
                }

                if (!clienteSelecionado) {
                    return alert('Por favor, selecione um cliente de destino');
                }

                const valorTransferencia = parseFloat(desformatarMoeda($('#valorTransferencia').val()));
                if (valorTransferencia <= 0) {
                    return alert('O valor de transferência desejado deve ser maior que zero');
                }

                if (valorTransferencia > preview.saldoOrigem) {
                    return alert('O valor da transferência não pode ser maior que o saldo disponível (R$ ' + formatarMoeda(preview.saldoOrigem) + ')');
                }

                salvarFormulario();
            });
        });
    </script>
</head>
<body>
    <div class="container">
        <div id="mensagemSucesso" style="display: none;">
            <div class="mensagem-sucesso">
                <i class="fas fa-check-circle"></i>
                <h4>Transferência criada com sucesso!</h4>
                <p>A solicitação de transferência de limite foi enviada para aprovação</p>
                <div class="mt-3"><a href="TransferenciaCreditoFormulario.asp">Clique aqui para criar uma nova transferência</a></div>
            </div>
        </div>

        <div id="containerFormulario">
            <div class="d-flex my-4">
                <h3>Nova Transferência de Limite de Crédito</h3>
                <button onclick="window.open('/corporativo/manuais/manual_solicitar_transferencia.pdf', '_blank')" class="btn-manual ms-4" title="Manual de Uso">
                    <i class="fas fa-file-alt"></i>
                </button>
            </div>

            <form name="formularioTransferencia" id="formularioTransferencia" method="post">
                <input type="hidden" id="hiddenCodClienteDestino" />
                <input type="hidden" id="hiddenCodLojaDestino" />
                                     
                <div class="card">
                    <div class="row g-3">
                        <div class="col-md-6">
                            <label for="plantaOrigem" class="form-label">Planta de Origem <span class="required-field">*</span></label>
                            <select class="form-select form-select-sm" name="plantaOrigem" id="plantaOrigem" required onchange="plantaAlterada()">
                                <option value="">Selecione a planta</option>
                                <option value="06">Maranguape</option>
                                <option value="00">Pirapetinga</option>
                                <option value="04">Uberaba</option>
                                <option value="05">Bento</option>
                                <option value="07">Saquarema</option>
                            </select>
                        </div>
                        <div class="col-md-6">
                            <label for="plantaDestino" class="form-label">Planta de Destino <span class="required-field">*</span></label>
                            <select class="form-select form-select-sm" name="plantaDestino" id="plantaDestino" required onchange="plantaAlterada()">
                                <option value="">Selecione a planta</option>
                                <option value="06">Maranguape</option>
                                <option value="00">Pirapetinga</option>
                                <option value="04">Uberaba</option>
                                <option value="05">Bento</option>
                                <option value="07">Saquarema</option>
                            </select>
                        </div>
                        <div class="col-md-10">
                            <label for="codClienteOrigem" class="form-label">Cód. Cliente (Origem)<span class="required-field me-2">*</span></label>
                            <div class="input-group">
                                <span class="btn btn-sm btn-secondary" onclick="PesquisaCliente()">Buscar <i class="fa fa-external-link-alt"></i></span>
                                <input onblur="obterDadosTransferencia()" type="text" class="form-control form-control-sm" id="codClienteOrigem" name="codClienteOrigem" required>
                            </div>
                            <small class="text-muted">Clique no botão para buscar ou digite um cód.</small>
                        </div>
                        <div class="col-md-2" id="divLoja">
                            <label for="codLojaOrigem" class="form-label">Cód. Loja (Origem) <span class="required-field me-2">*</span></label>
                            <input value="01" onblur="obterDadosTransferencia()" type="text" class="form-control form-control-sm" name="codLojaOrigem" id="codLojaOrigem" size="2" maxlength="2"/>
                        </div>
                    </div>

                    <div id="divLoadingCliente" style="display: none;">
                        <div class="divider"></div>
                        <div class="loading-container">
                            <div class="spinner-border"></div>
                            <p>Carregando dados do cliente...</p>
                        </div>
                    </div>

                    <div id="divErroCliente" style="display: none;">
                        <div class="divider"></div>
                        <div class="mensagem-erro">
                            <i class="fas fa-exclamation-triangle"></i>
                            <h6 id="mensagemErroCliente"></h6>
                        </div>
                    </div>

                    <div id="divDadosCliente" style="display: none;">
                        <div class="divider"></div>

                        <div class="mb-3">
                            <label for="nomeCliente" class="form-label">Cliente</label>
                            <input type="text" 
                                class="form-control form-control-sm" 
                                id="nomeCliente" 
                                value="Carregando..." 
                                disabled>
                        </div>

                        <div class="mb-3">
                            <label for="valorTransferencia" class="form-label">Valor a transferir <span class="required-field">*</span></label>
                            <input type="text" class="form-control form-control-sm" name="valorTransferencia" id="valorTransferencia" required />
                        </div>

                        <div class="transfer-container">
                            <div class="transfer-card">
                                <div class="transfer-header">
                                    <div>
                                        <div class="transfer-label">Origem</div>
                                        <div class="transfer-title" id="previewOriginName">-</div>
                                    </div>
                                    <div style="text-align: right;">
                                        <div class="text-muted">Cód. Cliente: <span id="previewOriginCode">-</span></div>
                                        <div class="text-muted">Cód. Loja: <span id="previewOriginLoja">-</span></div>
                                    </div>
                                </div>
                                <div class="transfer-details">
                                    <div class="transfer-row">
                                        <span class="transfer-row-label">Limite Atual</span>
                                        <span class="transfer-row-value" id="limiteAtualOrigem">-</span>
                                    </div>
                                    <div class="transfer-row">
                                        <span class="transfer-row-label">Em Aberto</span>
                                        <span class="transfer-row-value" id="limiteAbertoOrigem">-</span>
                                    </div>
                                    <div class="transfer-row">
                                        <span class="transfer-row-label">Saldo Disponível</span>
                                        <span class="transfer-row-value" id="saldoOrigem">-</span>
                                    </div>
                                    <div class="transfer-row highlight">
                                        <span class="transfer-row-label">Após Transferência</span>
                                        <span class="transfer-row-value" id="previewOriginAfter">R$ 0,00</span>
                                    </div>
                                </div>
                            </div>

                            <div class="transfer-arrow-container">
                                <div class="transfer-arrow">
                                    <i class="fas fa-arrow-right"></i>
                                </div>
                            </div>

                            <div class="transfer-card">
                                <div class="transfer-header">
                                    <div>
                                        <div class="transfer-label">Destino</div>
                                        <div class="transfer-title" id="previewDestName">-</div>
                                    </div>
                                    <div style="text-align: right;">
                                        <div class="text-muted">Cód. Cliente: <span id="previewDestCode">-</span></div>
                                        <div class="text-muted">Cód. Loja: <span id="previewDestLoja">-</span></div>
                                    </div>
                                </div>
                                <div class="transfer-details">
                                    <div class="transfer-row">
                                        <span class="transfer-row-label">Limite Atual</span>
                                        <span class="transfer-row-value" id="limiteAtualDestino">-</span>
                                    </div>
                                    <div class="transfer-row">
                                        <span class="transfer-row-label">Em Aberto</span>
                                        <span class="transfer-row-value" id="limiteAbertoDestino">-</span>
                                    </div>
                                    <div class="transfer-row">
                                        <span class="transfer-row-label">Saldo Disponível</span>
                                        <span class="transfer-row-value" id="saldoDestino">-</span>
                                    </div>
                                    <div class="transfer-row highlight">
                                        <span class="transfer-row-label">Após Transferência</span>
                                        <span class="transfer-row-value" id="previewDestAfter">R$ 0,00</span>
                                    </div>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>

                <div class="d-flex justify-content-end gap-2 mt-4 mb-4">
                    <button type="submit" class="btn btn-sm btn-primary" id="btnSalvar">
                        <span class="spinner-border spinner-border-sm me-2" role="status" aria-hidden="true" id="iconLoading" style="display: none;"></span>
                        <i class="fa fa-check me-2" id="iconCheck"></i> Salvar
                    </button>
                </div>
            </form>
        </div>
    </div>

    <div class="modal fade" id="modalSelecaoDestino" tabindex="-1" aria-hidden="true">
        <div class="modal-dialog modal-xl">
            <div class="modal-content">
                <div class="modal-header">
                    <h5 class="modal-title" id="modalSelecaoDestinoLabel">
                        <i class="fas fa-users me-2"></i>Selecione o Cliente de Destino
                    </h5>
                    <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Close"></button>
                </div>
                <div class="modal-body">
                    <p class="text-muted mb-3">Foram encontrados múltiplos clientes de destino. Por favor, selecione um dos registros abaixo:</p>
                    <div id="listaClientesDestino"></div>
                </div>
                <div class="modal-footer">
                    <button type="button" class="btn btn-sm btn-secondary" data-bs-dismiss="modal">Cancelar</button>
                    <button type="button" class="btn btn-sm btn-primary" id="btnConfirmarDestino" onclick="confirmarSelecaoDestino()" disabled>
                        <i class="fa fa-check me-2"></i>Confirmar Seleção
                    </button>
                </div>
            </div>
        </div>
    </div>
</body>
</html>