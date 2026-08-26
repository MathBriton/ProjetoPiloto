<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#Include File="../../../cnn2/cnn2.asp"-->
<!--#Include File="../../../funcoes/funcoes.asp"-->
<!--#Include File="../../../includes/variaveis.asp"-->
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

    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.0.2/dist/css/bootstrap.min.css" rel="stylesheet" integrity="sha384-EVSTQN3/azprG1Anm3QDgpJLIm9Nao0Yz1ztcQTwFspd3yD65VohhpuuCOmLASjC" crossorigin="anonymous" />
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

        .page-subtitle {
            color: #6c757d;
            font-size: .85rem;
        }

        .page-actions {
            gap: .5rem;
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

        .form-layout {
            display: grid;
            grid-template-columns: 270px minmax(0, 1fr);
        }

        .tabs-container {
            background-color: #f8f9fa;
            border-right: 1px solid #e9ecef;
            overflow-y: auto;
        }

        .tabs {
            display: flex;
            flex-direction: column;
            gap: 3px;
            padding: 14px 10px;
        }

        .tab-group-title {
            color: #9aa3af;
            font-size: 11px;
            font-weight: 600;
            letter-spacing: .06em;
            margin: 14px 10px 6px;
            text-transform: uppercase;
        }

        .tab-group-title:first-child {
            margin-top: 0;
        }

        .tab-button {
            align-items: center;
            background: transparent;
            border: none;
            border-radius: 7px;
            color: #495057;
            cursor: pointer;
            display: flex;
            font-weight: 500;
            gap: 9px;
            justify-content: flex-start;
            padding: 9px 11px;
            text-align: left;
            transition: all .2s ease;
            white-space: nowrap;
            width: 100%;
        }

        .tab-button:hover {
            background-color: #eef2f7;
        }

        .tab-button.active {
            background-color: #e7f1ff;
            color: #0d6efd;
            font-weight: 700;
        }

        .tab-button.completed {
            color: #198754;
        }

        .tab-button.invalid {
            background-color: #fff5f5;
            color: #dc3545;
        }

        .tab-button:disabled {
            cursor: not-allowed;
            opacity: .5;
        }

        #btnVoltarPagina:disabled,
        #btnProximaPagina:disabled {
            cursor: not-allowed;
            pointer-events: auto;
        }

        .tab-icon {
            flex-shrink: 0;
            font-size: 16px;
            width: 18px;
        }

        .tab-title {
            flex: 1;
            font-size: 14px;
        }

        .tab-meta {
            background: #e9ecef;
            border-radius: 999px;
            color: #6c757d;
            font-size: 10px;
            font-weight: 700;
            margin-left: auto;
            padding: 1px 7px;
        }

        .tab-button.active .tab-meta {
            background: #d7eaff;
            color: #0d6efd;
        }

        .tab-content-wrapper {
            background: #fff;
            min-width: 0;
        }

        .tab-content {
            animation: fadeIn .25s ease;
            display: none;
            padding: 30px;
            height: 80vh;
            overflow: auto;
        }

        .tab-content.active {
            display: block;
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

        .z-index-9999 {
            z-index: 9999 !important;
        }
        
        #modalRoteiroProduto {
            z-index: 1070;
        }
        .modal-backdrop.roteiro-produto-backdrop {
            z-index: 1065;
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
        
        .nav-footer {
            position: sticky;
            bottom: 0;
            background: #fff;
            border-top: 1px solid #dee2e6;
            padding: 10px 20px;
            display: flex;
            justify-content: flex-end;
            gap: 10px;
            z-index: 10;
        }

        table {
            width: max-content !important;
            min-width: 100%;
        }

        td, th {
            padding: 8px 12px !important;
        }

        th {
            font-weight: 600;
            color: #495057;
        }

        tr:hover {
            background: #f5f5f5;
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
      
        .switch-item {
            display: flex;
            align-items: self-end;
            gap: 1rem;
        }
        .switch-item .form-check-input {
            width: 2.75rem;
            height: 1.45rem;
            cursor: pointer;
            flex-shrink: 0;
        }
        .switch-label {
            font-size: .92rem;
            font-weight: 500;
            margin-bottom: .1rem;
            cursor: pointer;
        }
        .form-check-input {
            width: 1.2rem;
            height: 1.2rem;
            cursor: pointer;
        }

        #gradeItem th:last-child,
        #gradeItem td:last-child {
            position: sticky;
            right: 0;
            z-index: 5;
            white-space: nowrap;
        }
        
        #gradeItem td:last-child {
            background: linear-gradient(to right, rgba(0,0,0,0.1) 0%, transparent 1px),#fff;
            background-origin: border-box;
            background-clip: padding-box;
        }
        
        #gradeItem thead th:last-child {
            background: linear-gradient(to right, rgba(0,0,0,0.1) 0%, transparent 1px), #f5f5f5;
            background-origin: border-box;
            background-clip: padding-box;
            z-index: 10;
        }

        #gradeItem tfoot td:last-child {
            background: linear-gradient(to right, rgba(0,0,0,0.1) 0%, transparent 1px), #f5f5f5;
            background-origin: border-box;
            background-clip: padding-box;
            z-index: 10;
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

        @media (max-width: 1400px) {
            .container {
                max-width: 95%;
            }

            .form-layout {
                grid-template-columns: 1fr;
            }

            .tabs-container {
                border-bottom: 1px solid #e9ecef;
                border-right: none;
                overflow-x: auto;
            }

            .tabs {
                flex-direction: row;
                min-width: max-content;
                padding: 8px;
            }

            .tab-group-title,
            .tab-meta {
                display: none;
            }

            .tab-button {
                justify-content: center;
                min-width: 135px;
                text-align: center;
                width: auto;
            }

            .tab-title {
                flex: initial;
            }
        }
    </style>

    <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.0.2/dist/js/bootstrap.bundle.min.js" integrity="sha384-MrcW6ZMFYlzcLA8Nl+NtUVF0sA7MsXsP1UyJoMp4YLEuNSfAP+JcXn/tWtIaxVXM" crossorigin="anonymous"></script>
    <script src="https://cdn.jsdelivr.net/npm/select2@4.1.0-rc.0/dist/js/select2.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/jquery.mask/1.14.15/jquery.mask.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/just-validate@latest/dist/just-validate.production.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/jquery-maskmoney/3.0.2/jquery.maskMoney.min.js"></script>

    <script src="/corporativo/funcoes/funcoesJAVASCRIPT.js?v=<%=Now()%>"></script>
    <script src="/corporativo/components/toast.js?v=<%=Now()%>"></script>
    <script src="/corporativo/components/confirmation.js?v=<%=Now()%>"></script>
    <script src="/corporativo/components/upload/fileUpload.js?v=<%=Now()%>"></script>
    <script src="/corporativo/components/dropdownMenu.js?v=<%=Now()%>"></script>

    <script>
        let id = 0;
        let abaAtual = 1;
        let totalAbas = 2;
        let abasValidadas = new Set();
        let validacoes = {};

        // Modal de Item
        let listaItens = [];
        let itemEditandoIndex = null;
        let resultadosTes = [];
        let tesSelecionado = null;
        let resultadosProduto = [];
        let produtoSelecionado = null;
        let roteiroProdutoSelecionado = [];
        let resultadosClienteRemessa = [];
        let clienteRemessaSelecionado = null;

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

        function inicializarMascaras() {
            $('#itemQuantidade, #itemValorTotal, #itemAreaTotal').mask('000.000.000.000.000', { reverse: true });
           
            $('#itemValorUnitario').maskMoney({
                prefix: '',
                thousands: '.',
                decimal: ',',
                allowZero: true,
                allowNegative: false
            });
        }

        function carregarTiposItem(codigoPlanta) {
            const $select = $('#itemTipo');
            $select.empty().append('<option value="">Selecione...</option>');

            if (!codigoPlanta) return;

            return $.ajax({
                url: 'EmissaoPedidoRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: { acao: 'listarTiposItem', codigoPlanta: codigoPlanta },
                dataType: 'json',
                success: function(response) {
                    if (!response.sucesso) {
                        console.error('Erro ao carregar tipos de item: ' + response.mensagem);
                        return;
                    }

                    (response.dados || []).forEach(function(item) {
                        $select.append($('<option>', {
                            value: item.codigo,
                            text: item.descricao
                        }));
                    });
                },
                error: function(xhr, status, error) {
                    console.error('Erro ao carregar tipos de item:', error);
                }
            });
        }

        function desformatarMoeda(valor) {
            const texto = String(valor || '0');
            const numero = parseFloat(texto.replaceAll('.', '').replaceAll(',', '.'));
            return isNaN(numero) ? 0 : numero;
        }

        function plantaAtual() {
            return $('#codigoPlanta').val();
        }

        function atualizarCamposPorTipoItem() {
            const tipoItem = $('#itemTipo').val();
            const camposChapa = '#itemLargura, #itemComprimento, #itemVinco1, #itemVinco2, #itemVinco3, #itemVinco4, #itemVinco5';
            const camposPapel = '#itemGramatura, #itemDiametroBobina, #itemLarguraBobina, #itemTubete';

            $(camposChapa).prop('disabled', tipoItem !== 'CH');
            $(camposPapel).prop('disabled', tipoItem !== 'PA');
            $('#itemGramatura, #itemTubete').trigger('change.select2');
        }

        function abrirModalLimiteCliente() {
            const codigoCliente = $('#codigoCliente').val();
            const codigoLoja = $('#codigoLoja').val();

            if (!codigoCliente || !codigoLoja) {
                Toast.warning('Selecione o cliente antes de visualizar o limite de crédito.');
                return;
            }

            $('#limiteClienteConteudo').html(`
                <div class="text-center py-4">
                    <div class="spinner-border spinner-border-sm text-primary" role="status">
                        <span class="visually-hidden">Carregando...</span>
                    </div>
                    <p class="mt-2 mb-0 text-muted">Consultando informações financeiras...</p>
                </div>
            `);

            const modal = new bootstrap.Modal(document.getElementById('modalLimiteCliente'));
            modal.show();

            $.ajax({
                url: 'EmissaoPedidoRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: {
                    acao: 'obterSaldoCliente',
                    codigoPlanta: $('#codigoPlanta').val(),
                    codigoCliente: codigoCliente,
                    codigoLoja: codigoLoja
                },
                dataType: 'json',
                success: function(response) {
                    if (response.sucesso === false) {
                        $('#limiteClienteConteudo').html(`<div class="alert alert-danger mb-0" role="alert">${response.mensagem || 'Erro ao consultar o limite de crédito do cliente.'}</div>`);
                        return;
                    }

                    const dados = response.dados || response;
                    $('#limiteClienteConteudo').html(`
                        <div class="table-responsive">
                            <table class="table table-sm table-bordered mb-0">
                                <tbody>
                                    <tr><th class="table-light">Limite</th><td class="text-end">${formatarMoeda((dados.limite))}</td></tr>
                                    <tr><th class="table-light">Em aberto</th><td class="text-end">${formatarMoeda((dados.aberto))}</td></tr>
                                    <tr><th class="table-light">Pedidos</th><td class="text-end">${formatarMoeda((dados.pedidos))}</td></tr>
                                    <tr><th class="table-light">Pedidos a liberar</th><td class="text-end">${formatarMoeda((dados.pedidosLiberar))}</td></tr>
                                    <tr><th class="table-light">Saldo</th><td class="text-end fw-bold">${formatarMoeda((dados.saldo))}</td></tr>
                                </tbody>
                            </table>
                        </div>
                    `);
                },
                error: function(xhr, status, error) {
                    console.error('Erro ao consultar limite de crédito:', error);
                    $('#limiteClienteConteudo').html('<div class="alert alert-danger mb-0" role="alert">Não foi possível consultar o limite de crédito do cliente.</div>');
                }
            });
        }

        function exibirCampoPorPlanta() {
            const planta = plantaAtual();
            const exibeCamposBento = planta === '05';

            $('.campos-bento').toggleClass('d-none', !exibeCamposBento);
            $('.campos-bento').find('input, select, textarea').prop('disabled', !exibeCamposBento);
            $('.campos-bento select').trigger('change.select2');
        }

        function definirMarcacaoObrigatoria(campoId, obrigatorio) {
            const $label = $(`label[for="${campoId}"]`);
            if (!$label.length) return;

            let $span = $label.find('.required-field');
            if (obrigatorio) {
                if (!$span.length) {
                    $label.append(' <span class="required-field">*</span>');
                } else {
                    $span.show();
                }
            } else if ($span.length) {
                $span.hide();
            }
        }

        function atualizarMarcacoesObrigatorias() {
            const tesRemessa = $('#hidTesRemessa').val() === 'S';
            const poder3 = $('#hidPoder3').val() === 'S';
            const clienteRemessaObrigatorio = tesRemessa && !poder3;

            definirMarcacaoObrigatoria('codigoClienteRemessa', clienteRemessaObrigatorio);
            definirMarcacaoObrigatoria('codigoLojaRemessa', clienteRemessaObrigatorio);
            definirMarcacaoObrigatoria('nomeClienteRemessa', clienteRemessaObrigatorio);
            definirMarcacaoObrigatoria('saldoAparas', poder3);
        }

        function exibirCamposCondicionais() {
            const tesRemessa = $('#hidTesRemessa').val() === 'S';
            const poder3 = $('#hidPoder3').val() === 'S';
            const exibeCamposRemessa = tesRemessa && !poder3;

            $('.campos-remessa').toggleClass('d-none', !exibeCamposRemessa);
            $('.campo-saldo-aparas').toggleClass('d-none', !poder3);
            atualizarMarcacoesObrigatorias();
        }

        function calcularTotaisItem() {
            const quantidade = desformatarMoeda($('#itemQuantidade').val());
            const valorUnitario = desformatarMoeda($('#itemValorUnitario').val());
            const areaRefile = $('#itemAreaRefile').val();
            $('#itemValorTotal').val(formatarMoeda(quantidade * valorUnitario));
            $('#itemAreaTotal').val(formatarMoeda(quantidade * areaRefile));
        }

        function validarItens() {
            const itemAtual = {
                lotePiloto: $('#lotePiloto').is(':checked') ? 1 : 0,
                itemTipo: $('#itemTipo').val(),
                codigoProduto: $('#itemCodigoProduto').val(),
                referencia: $('#itemReferencia').val(),
                estrutura: $('#itemEstrutura').val(),
                quantidade: desformatarMoeda($('#itemQuantidade').val()),
                valorUnitario: desformatarMoeda($('#itemValorUnitario').val()),
                largura: $('#itemLargura').val(),
                vinco1: $('#itemVinco1').val(),
                vinco2: $('#itemVinco2').val(),
                vinco3: $('#itemVinco3').val(),
                vinco4: $('#itemVinco4').val(),
                vinco5: $('#itemVinco5').val(),
                comprimento: $('#itemComprimento').val(),
                gramatura: $('#itemGramatura').val(),
                areaRefile: $('#itemAreaRefile').val(),
                area: $('#itemArea').val(),
                peso: $('#itemPeso').val(),
                poder3: $('#hidPoder3').val(),
                diametroBobina: $('#itemDiametroBobina').val(),
                larguraBobina: $('#itemLarguraBobina').val(),
                tubeteBobina: $('#itemTubete').val(),
                modalFrete: $('#itemModalFrete').val(),
                comercioExterior: $('#itemComercioExterior').val(),
                dataEntrega: $('#itemDataEntrega').val(),
                produtoEstoque: $('#itemProdutoEstoque').val()
            }


            // Filtra itens removendo item atual
            const itens = [
                ...listaItens.filter((item, index) => index != itemEditandoIndex),
                itemAtual
            ]
            

            const dados = {
                acao: 'validarItens',
                codigoPlanta: $('#codigoPlanta').val(),
                codigoCliente: $('#codigoCliente').val(),
                codigoLoja: $('#codigoLoja').val(),
                razaoSocial: $('#razaoSocial').val(),
                tes: $('#tes').val(),
                condicaoPagamento: $('#condicaoPagamento').val(),
                moeda: $('#moeda').val(),
                tipoFrete: $('#tipoFrete').val(),
                codigoClienteRemessa: $('#codigoClienteRemessa').val(),
                codigoLojaRemessa: $('#codigoLojaRemessa').val(),
                validaLimiteCredito: $('#hidValidaLimiteCredito').val(),
                toleranciaMaior: $('#hidToleranciaMaior').val(),
                toleranciaDataMaior: $('#hidToleranciaDataMaior').val(),
                tesRemessa: $('#hidTesRemessa').val(),
                itens: JSON.stringify(itens.map(({ roteiroProduto, ...item }) => item))
            };

            return $.ajax({
                url: 'EmissaoPedidoRepositorio.asp',
                type: 'POST',
                dataType: 'json',
                data: dados
            });
        }

        function VerificarExisteLotePiloto() {
            return $.ajax({
                url: 'EmissaoPedidoRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: {
                    acao: 'verificarExisteLotePiloto',
                    codigoPlanta: $('#codigoPlanta').val(),
                    codigoCliente: $('#codigoCliente').val(),
                    codigoLoja: $('#codigoLoja').val(),
                    codigoProduto: $('#itemCodigoProduto').val(),
                },
                dataType: 'json'
            });
        }

        function existeItemLotePilotoDuplicado() {
            const codigoProduto = $('#itemCodigoProduto').val();
            return listaItens.some(function(item, index) {
                return index !== itemEditandoIndex
                    && item.codigoProduto === codigoProduto
                    && item.lotePiloto == 1;
            });
        }

        function verificarExisteOrcamento() {
            return $.ajax({
                url: 'EmissaoPedidoRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: {
                    acao: 'verificarExisteOrcamento',
                    codigoPlanta: $('#codigoPlanta').val(),
                    codigoCliente: $('#codigoCliente').val(),
                    codigoLoja: $('#codigoLoja').val(),
                    codigoProduto: $('#itemCodigoProduto').val(),
                    tipoFrete: $('#tipoFrete').val(),
                    condicaoPagamento: $('#condicaoPagamento').val(),
                    valorUnitario: desformatarMoeda($('#itemValorUnitario').val())
                },
                dataType: 'json'
            });
        }

        function verificarPedidoCompraDuplicado() {
            return $.ajax({
                url: 'EmissaoPedidoRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: {
                    acao: 'verificarPedidoCompraDuplicado',
                    codigoPlanta: $('#codigoPlanta').val(),
                    codigoCliente: $('#codigoCliente').val(),
                    codigoProduto: $('#itemCodigoProduto').val(),
                    pedidoCompra: $('#itemPedidoCompra').val(),
                    itemPedido: $('#itemPedido').val()
                },
                dataType: 'json'
            });
        }

        function inicializarValidacoes() {
            const naoVazio = (msg) => ({
                validator: (value) => value !== null && value !== undefined && String(value).trim() !== '',
                errorMessage: msg
            });
            const maiorQueZero = (msg) => ({
                validator: (value) => value !== null && value !== undefined && parseFloat(value) > 0,
                errorMessage: msg
            });
            const obrigatorioQuando = (condicao, msg) => ({
                validator: (value) => !condicao() || (value !== null && value !== undefined && String(value).trim() !== '' && value != '0'),
                errorMessage: msg
            });

            validacoes.aba1 = new JustValidate('#formularioPedido', { lockForm: false });
            validacoes.aba1
                .addField('#codigoPlanta', [naoVazio('Planta é obrigatória')])
                .addField('#codigoCliente', [naoVazio('Cliente é obrigatório')])
                .addField('#codigoLoja', [naoVazio('Loja é obrigatória')])
                .addField('#razaoSocial', [naoVazio('Nome do cliente é obrigatório')])
                .addField('#tes', [naoVazio('TES é obrigatória')])
                .addField('#condicaoPagamento', [naoVazio('Condição de pagamento é obrigatória')])
                .addField('#moeda', [naoVazio('Moeda é obrigatória')])
                .addField('#tipoFrete', [naoVazio('Tipo de Frete é obrigatório')]);

            validacoes.aba2 = new JustValidate('#formularioPedido', { lockForm: false });
            validacoes.modalItem = new JustValidate('#formularioItem', { lockForm: false });
            
            const itemPapel = () => $('#itemTipo').val() === 'PA';
            const itemNaoPapel = () => $('#itemTipo').val() !== 'PA';
            const itemChapa = () => $('#itemTipo').val() === 'CH';

            validacoes.modalItem
                .addField('#itemTipo', [naoVazio('Tipo do item é obrigatório')])
                .addField('#itemCodigoProduto', [naoVazio('Produto é obrigatório')])
                .addField('#itemReferencia', [naoVazio('Referência é obrigatória')])
                .addField('#itemEstrutura', [obrigatorioQuando(itemNaoPapel, 'Estrutura é obrigatória')])
                .addField('#itemQuantidade', [maiorQueZero('Quantidade deve ser maior que zero')])
                .addField('#itemValorUnitario', [maiorQueZero('Valor unitário deve ser maior que zero')])
                .addField('#itemValorTotal', [maiorQueZero('Total deve ser maior que zero')])
                .addField('#itemAreaTotal', [obrigatorioQuando(itemNaoPapel, 'Área total deve ser maior que zero')])
                .addField('#itemDataEntrega', [naoVazio('Data de entrega é obrigatória')])
                .addField('#itemGramatura', [obrigatorioQuando(itemPapel, 'Gramatura é obrigatória')])
                .addField('#itemDiametroBobina', [obrigatorioQuando(itemPapel, 'Diâmetro da bobina é obrigatório')])
                .addField('#itemLarguraBobina', [obrigatorioQuando(itemPapel, 'Largura da bobina é obrigatória')])
                .addField('#itemTubete', [obrigatorioQuando(itemPapel, 'Tubete da bobina é obrigatório')])
                .addField('#itemLargura', [obrigatorioQuando(itemChapa, 'Largura é obrigatória')])
                .addField('#itemComprimento', [obrigatorioQuando(itemChapa, 'Comprimento é obrigatório')]);
        }

        function navegarParaGrade() {
            window.location.href = `../pedido_digitado/Pedidodigitado.asp?Planta=${$('#codigoPlanta').val()}`
        }

        function navegarParaAba(numeroAba) {
            if (numeroAba < 1 || numeroAba > totalAbas) return;

            abaAtual = numeroAba;

            $('.tab-button').removeClass('active');
            $(`.tab-button[data-tab="${numeroAba}"]`).addClass('active');

            $('.tab-content').removeClass('active');
            $(`.tab-content[data-tab-content="${numeroAba}"]`).addClass('active');

            atualizarBotoesNavegacao();
        }

        function atualizarBotoesNavegacao() {
            $('#btnVoltarPagina').prop('disabled', abaAtual <= 1);
            $('#btnProximaPagina').prop('disabled', abaAtual >= totalAbas);
        }

        function marcarAbaComoCompleta(numeroAba) {
            const $button = $(`.tab-button[data-tab="${numeroAba}"]`);

            if ($button.length) {
                $button.addClass('completed').removeClass('invalid');
                const $icon = $button.find('.tab-icon');
                if ($icon.length) {
                    $icon.attr('class', 'tab-icon bi bi-check-circle-fill');
                }
            }

            abasValidadas.add(numeroAba);
        }

        function marcarAbaComoInvalida(numeroAba) {
            const $button = $(`.tab-button[data-tab="${numeroAba}"]`);
            if ($button.length && !$button.hasClass('completed')) {
                $button.addClass('invalid');
                const $icon = $button.find('.tab-icon');
                if ($icon.length) {
                    $icon.attr('class', 'tab-icon bi bi-exclamation-circle-fill');
                }
            }
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
                url: 'EmissaoPedidoRepositorio.asp?v=<%=Now()%>',
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
                        <td>${cliente.razaoSocial}</td>
                        <td>${cliente.codigoCliente}</td>
                        <td>${cliente.codigoLoja}</td>
                        <td>${cliente.tipoPessoa == 'F' ? formatarCPF(cliente.cpfCnpj) : formatarCNPJ(cliente.cpfCnpj)}</td>
                        <td>${cliente.municipio}</td>
                        <td>${cliente.uf}</td>
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
            $('#tes').val('');
            
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
        
        function limparCamposAposCliente() {
            $('#tes, #codigoClienteRemessa, #codigoLojaRemessa, #nomeClienteRemessa').val('');
            $('#saldoAparas').val('0');
            $('#condicaoPagamento').empty().append('<option value="">Selecione a condição de pagamento</option>').trigger('change.select2');
            $('#moeda, #tipoFrete').val('').trigger('change.select2');
            $('#hidTesRemessa, #hidPoder3, #hidValidaLimiteCredito').val('N');
            
            listaItens = [];
            renderizarListaItens();
            exibirCamposCondicionais();
        }

        function limparCamposAposPlanta() {
            $('#codigoCliente, #codigoLoja, #razaoSocial').val('');
            $('#hidToleranciaMaior, #hidToleranciaDataMaior').val('0');
            limparCamposAposCliente();
        }

        function preencherDadosCliente(dados) {
            limparCamposAposCliente();
            $('#razaoSocial').val(dados.razaoSocial || '');
            $('#codigoCliente').val(dados.codigoCliente || '');
            $('#codigoLoja').val(dados.codigoLoja || '');
            $('#hidToleranciaMaior').val(dados.toleranciaMaior ||  0)
            $('#hidToleranciaDataMaior').val(dados.toleranciaDataMaior ||  0)
            $('#saldoAparas').val(dados.saldoAparas || 0)
            $('#tipoFrete').val(dados.tipoFrete || 0).trigger('change.select2')
            $('#moeda').val('0').trigger('change.select2')

            carregarCondicoesPagamento($('#codigoPlanta').val(), dados.condicaoPagamento);
        }

        function abrirModalClienteRemessa() {
            const codigoPlanta = $('#codigoPlanta').val();

            if (!codigoPlanta) { 
                Toast.warning('Informe o código da planta antes de buscar por clientes remessa.'); 
                return; 
            }

            clienteRemessaSelecionado = null;
            resultadosClienteRemessa = [];

            $('#clienteRemessaBusca').val('');
            $('#clienteRemessaResultados').html(`
                <div class="text-center text-muted">
                    <i class="bi bi-search" style="font-size: 1.5rem;"></i>
                    <p class="mt-2 mb-0">Digite um termo para buscar clientes remessa existentes.</p>
                </div>
            `);

            $('#btnConfirmarClienteRemessa').prop('disabled', true);

            const modal = new bootstrap.Modal(document.getElementById('modalClienteRemessa'));
            modal.show();

            setTimeout(function() { 
                $('#clienteRemessaBusca').focus(); 
            }, 500);
        }

        function pesquisarClienteRemessa() {
            const busca = $('#clienteRemessaBusca').val().trim();
            const codigoPlanta = $('#codigoPlanta').val();
            const codigoCliente = $('#codigoCliente').val();
            const codigoLoja = $('#codigoLoja').val();

            if (!busca) {
                $('#clienteRemessaResultados').html(`
                    <div class="text-center text-muted">
                        <i class="bi bi-search" style="font-size: 1.5rem;"></i>
                        <p class="mt-2 mb-0">Digite um termo para buscar clientes remessa existentes.</p>
                    </div>
                `);
                return;
            }

            clienteRemessaSelecionado = null;
            $('#btnConfirmarClienteRemessa').prop('disabled', true);

            $('#clienteRemessaResultados').html(`
                <div class="text-center">
                    <div class="spinner-border spinner-border-sm text-primary" role="status">
                        <span class="visually-hidden">Buscando...</span>
                    </div>
                    <p class="mt-2 mb-0 text-muted">Buscando clientes remessa...</p>
                </div>
            `);

            $.ajax({
                url: 'EmissaoPedidoRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: {
                    acao: 'listarClientesRemessa',
                    codigoPlanta: codigoPlanta,
                    codigoCliente: codigoCliente,
                    codigoLoja: codigoLoja,
                    busca: busca
                },
                dataType: 'json',
                success: function(response) {
                    if (!response.sucesso) {
                        $('#clienteRemessaResultados').html(`
                            <div class="alert text-center alert-danger" role="alert">
                                <i class="bi bi-exclamation-triangle" style="font-size: 1.5rem;"></i>
                                <p class="mt-2 mb-0"> ${response.mensagem || 'Erro ao buscar clienteRemessas.'} </p>
                            </div>
                        `);
                        return;
                    }

                    resultadosClienteRemessa = response.dados || [];

                    if (resultadosClienteRemessa.length === 0) {
                        $('#clienteRemessaResultados').html(`
                            <div class="alert text-center alert-info" role="alert">
                                <i class="bi bi-info-circle" style="font-size: 1.5rem;"></i>
                                <p class="mt-2 mb-0"> Nenhum clienteRemessa encontrado com os termos pesquisados. </p>
                            </div>
                        `);
                        return;
                    }

                    renderizarResultadosClienteRemessa();
                },
                error: function(xhr, status, error) {
                    console.error('Erro ao buscar clienteRemessas:', error);
                    $('#clienteRemessaResultados').html(`
                        <div class="alert text-center alert-info" role="alert">
                            <i class="bi bi-info-circle" style="font-size: 1.5rem;"></i>
                            <p class="mt-2 mb-0"> Nenhum clienteRemessa encontrado com os termos pesquisados.</p>
                        </div>
                    `);
                }
            });
        }

        function renderizarResultadosClienteRemessa() {
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
                                <th>Municipio</th>
                                <th>UF</th>
                            </tr>
                        </thead>
                        <tbody>
            `;

            resultadosClienteRemessa.forEach(function(clienteRemessa, index) {
                html += `
                    <tr data-index="${index}">
                        <td class="text-center">
                            <button onclick="selecionarClienteRemessa(${index})"
                                class="btn-selecionar" 
                                title="Selecionar este clienteRemessa">
                                <i class="bi bi-circle"></i> 
                            </button>
                        </td>
                        <td>${clienteRemessa.razaoSocial}</td>
                        <td>${clienteRemessa.codigoCliente}</td>
                        <td>${clienteRemessa.codigoLoja}</td>
                        <td>${formatarCNPJ(clienteRemessa.cpfCnpj)}</td>
                        <td>${clienteRemessa.municipio}</td>
                        <td>${clienteRemessa.uf}</td>
                    </tr>
                `;
            });

            html += `
                        </tbody>
                    </table>
                </div>
            `;

            $('#clienteRemessaResultados').html(html);
        }

        function selecionarClienteRemessa(index) {
            clienteRemessaSelecionado = resultadosClienteRemessa[index];
            
            $('#clienteRemessaResultados .btn-selecionar')
                .removeClass('selecionado')
                .html('<i class="bi bi-circle"></i>')
                .attr('title', 'Selecionar este item');
            
            $('#clienteRemessaResultados tr[data-index="' + index + '"] .btn-selecionar')
                .addClass('selecionado')
                .html('<i class="bi bi-check-circle-fill"></i>')
                .attr('title', 'Item Selecionado');
            
            $('#btnConfirmarClienteRemessa').prop('disabled', false);
        }

        function confirmarClienteRemessa() {
            if (!clienteRemessaSelecionado) {
                Toast.warning('Selecione um clienteRemessa');
                return;
            }

            const modal = bootstrap.Modal.getInstance(document.getElementById('modalClienteRemessa'));
            modal.hide();

            preencherDadosClienteRemessa(clienteRemessaSelecionado);
        }
        
        function preencherDadosClienteRemessa(dados) {
            $('#codigoClienteRemessa').val(dados.codigoCliente || '');
            $('#codigoLojaRemessa').val(dados.codigoLoja || '');
            $('#nomeClienteRemessa').val(dados.razaoSocial || '');
        }

        function abrirModalTes() {
            const codigoPlanta = $('#codigoPlanta').val();
            const codigoCliente = $('#codigoCliente').val();
            const codigoLoja = $('#codigoLoja').val();

            if (!codigoPlanta || !codigoCliente || !codigoLoja) {
                Toast.warning('Informe planta e cliente antes de buscar TES.');
                return;
            }

            tesSelecionado = null;
            resultadosTes = [];

            $('#btnConfirmarTes').prop('disabled', true);
            $('#tesResultados').html(`
                <div class="text-center">
                    <div class="spinner-border spinner-border-sm text-primary" role="status">
                        <span class="visually-hidden">Buscando...</span>
                    </div>
                    <p class="mt-2 mb-0 text-muted">Buscando TES...</p>
                </div>
            `);

            const modal = new bootstrap.Modal(document.getElementById('modalTes'));
            modal.show();
            pesquisarTes();
        }

        function pesquisarTes() {
            const codigoPlanta = $('#codigoPlanta').val();
            const codigoCliente = $('#codigoCliente').val();
            const codigoLoja = $('#codigoLoja').val();

            $.ajax({
                url: 'EmissaoPedidoRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: {
                    acao: 'listarTes',
                    codigoPlanta: codigoPlanta,
                    codigoCliente: codigoCliente,
                    codigoLoja: codigoLoja
                },
                dataType: 'json',
                success: function(response) {
                    if (!response.sucesso) {
                        $('#tesResultados').html(`
                            <div class="alert text-center alert-danger" role="alert">
                                <i class="bi bi-exclamation-triangle" style="font-size: 1.5rem;"></i>
                                <p class="mt-2 mb-0"> ${response.mensagem || 'Erro ao buscar TES.'} </p>
                            </div>
                        `);
                        return;
                    }

                    resultadosTes = response.dados || [];

                    if (resultadosTes.length === 0) {
                        $('#tesResultados').html(`
                            <div class="alert text-center alert-info" role="alert">
                                <i class="bi bi-info-circle" style="font-size: 1.5rem;"></i>
                                <p class="mt-2 mb-0"> Nenhuma TES encontrada. </p>
                            </div>
                        `);
                        return;
                    }

                    renderizarResultadosTes();
                },
                error: function(xhr, status, error) {
                    console.error('Erro ao buscar TES:', error);
                    $('#tesResultados').html(`
                        <div class="alert text-center alert-info" role="alert">
                            <i class="bi bi-info-circle" style="font-size: 1.5rem;"></i>
                            <p class="mt-2 mb-0"> Nenhuma TES encontrada.</p>
                        </div>
                    `);
                }
            });
        }

        function renderizarResultadosTes() {
            let html = `
                <div class="table-responsive">
                    <table class="table table-sm table-hover mb-0">
                        <thead class="thead-light">
                            <tr>
                                <th style="width: 50px;"></th>
                                <th>TES</th>
                                <th>Descrição</th>
                            </tr>
                        </thead>
                        <tbody>
            `;

            resultadosTes.forEach(function(tes, index) {
                html += `
                    <tr data-index="${index}">
                        <td class="text-center">
                            <button onclick="selecionarTes(${index})"
                                class="btn-selecionar"
                                title="Selecionar esta TES">
                                <i class="bi bi-circle"></i>
                            </button>
                        </td>
                        <td>${tes.codigo}</td>
                        <td>${tes.descricao}</td>
                    </tr>
                `;
            });

            html += `
                        </tbody>
                    </table>
                </div>
            `;

            $('#tesResultados').html(html);
        }

        function selecionarTes(index) {
            tesSelecionado = resultadosTes[index];

            $('#tesResultados .btn-selecionar')
                .removeClass('selecionado')
                .html('<i class="bi bi-circle"></i>')
                .attr('title', 'Selecionar esta TES');

            $('#tesResultados tr[data-index="' + index + '"] .btn-selecionar')
                .addClass('selecionado')
                .html('<i class="bi bi-check-circle-fill"></i>')
                .attr('title', 'TES Selecionada');

            $('#btnConfirmarTes').prop('disabled', false);
        }

        function confirmarTes() {
            if (!tesSelecionado) {
                Toast.warning('Selecione uma TES');
                return;
            }

            const modal = bootstrap.Modal.getInstance(document.getElementById('modalTes'));
            modal.hide();
            preencherDadosTes(tesSelecionado);
        }

        function preencherDadosTes(dados) {
            $('#tes').val(dados.codigo);
            $('#hidTesRemessa').val(dados.tesRemessa);
            $('#hidPoder3').val(dados.poder3);
            $('#hidValidaLimiteCredito').val(dados.validaLimiteCredito);
            exibirCamposCondicionais();
        }

        function abrirModalProduto() {
            const codigoPlanta = $('#codigoPlanta').val();
            const codigoCliente = $('#codigoCliente').val();
            const codigoLoja = $('#codigoLoja').val();

            if (!codigoPlanta || !codigoCliente || !codigoLoja) {
                Toast.warning('Informe planta e cliente antes de buscar produtos.');
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

            const modal = new bootstrap.Modal(document.getElementById('modalProduto'));
            modal.show();
            setTimeout(function() { $('#produtoBusca').focus(); }, 500);
        }

        function pesquisarProduto() {
            const busca = $('#produtoBusca').val().trim();
            const codigoPlanta = $('#codigoPlanta').val();
            const codigoCliente = $('#codigoCliente').val();
            const codigoLoja = $('#codigoLoja').val();
            const itemTipo = $('#itemTipo').val();

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
                url: 'EmissaoPedidoRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: {
                    acao: 'listarProdutos',
                    itemTipo: itemTipo,
                    codigoPlanta: codigoPlanta,
                    codigoCliente: codigoCliente,
                    codigoLoja: codigoLoja,
                    busca: busca
                },
                dataType: 'json',
                success: function(response) {
                    if (!response.sucesso) {
                        $('#produtoResultados').html(`<div class="alert text-center alert-danger" role="alert"><p class="mb-0">${response.mensagem || 'Erro ao buscar produtos.'}</p></div>`);
                        return;
                    }

                    resultadosProduto = response.dados || [];

                    if (resultadosProduto.length === 0) {
                        $('#produtoResultados').html(`<div class="alert text-center alert-info" role="alert"><p class="mb-0">Nenhum produto encontrado com os termos pesquisados.</p></div>`);
                        return;
                    }

                    renderizarResultadosProduto();
                },
                error: function(xhr, status, error) {
                    console.error('Erro ao buscar produtos:', error);
                    $('#produtoResultados').html(`<div class="alert text-center alert-danger" role="alert"><p class="mb-0">Erro ao buscar produtos.</p></div>`);
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
                                <th>Descrição</th>
                                <th>Referência</th>
                                <th>Estrutura</th>
                            </tr>
                        </thead>
                        <tbody>
            `;
            resultadosProduto.forEach(function(produto, index) {
                html += `
                    <tr data-index="${index}">
                        <td class="text-center">
                            <button onclick="selecionarProduto(${index})" class="btn-selecionar" title="Selecionar este produto">
                                <i class="bi bi-circle"></i>
                            </button>
                        </td>
                        <td>${produto.codigo}</td>
                        <td>${produto.descricao || ''}</td>
                        <td>${produto.referencia || ''}</td>
                        <td>${produto.estrutura || ''}</td>
                    </tr>
                `;
            });
            html += `</tbody></table></div>`;
            $('#produtoResultados').html(html);
        }

        function selecionarProduto(index) {
            produtoSelecionado = resultadosProduto[index];

            $('#produtoResultados .btn-selecionar')
                .removeClass('selecionado')
                .html('<i class="bi bi-circle"></i>')
                .attr('title', 'Selecionar este produto');

            $('#produtoResultados tr[data-index="' + index + '"] .btn-selecionar')
                .addClass('selecionado')
                .html('<i class="bi bi-check-circle-fill"></i>')
                .attr('title', 'Produto Selecionado');

            $('#btnConfirmarProduto').prop('disabled', false);
        }

        async function confirmarProduto() {
            if (!produtoSelecionado) {
                Toast.warning('Selecione um produto');
                return;
            }

            const modal = bootstrap.Modal.getInstance(document.getElementById('modalProduto'));
            modal.hide();

            preencherDadosProduto(produtoSelecionado);
            exibirLoading(true);
            try {
                await carregarRoteiroProdutoSelecionado();
            } finally {
                exibirLoading(false);
            }
        }

        function carregarRoteiroProdutoSelecionado() {
            roteiroProdutoSelecionado = [];
            return $.ajax({
                url: 'EmissaoPedidoRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: {
                    acao: 'listarProdutoRoteiro',
                    codigoProduto: $('#itemCodigoProduto').val(),
                    codigoCliente: $('#codigoCliente').val(),
                    codigoLoja: $('#codigoLoja').val(),
                    tipoProduto: $('#itemProdutoTipo').val(),
                    codigoPlanta: $('#codigoPlanta').val()
                },
                dataType: 'json',
                success: function(response) {
                    if (!response.sucesso) {
                        Toast.error(response.mensagem || 'Erro ao buscar roteiro do produto.');
                        return;
                    }
                    roteiroProdutoSelecionado = response.dados || [];
                },
                error: function(xhr, status, error) {
                    console.error('Erro ao buscar roteiro do produto:', error);
                    Toast.error('Erro ao buscar roteiro do produto.');
                }
            });
        }

        function ObterValorUnitarioProdutoOrcamento() {
            const codigoProduto = $('#itemCodigoProduto').val();
            const valorUnitarioAtual = $('#itemValorUnitario').val();
            if (!codigoProduto) return;

            $.ajax({
                url: 'EmissaoPedidoRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: {
                    acao: 'obterValorUnitarioProdutoOrcamento',
                    codigoPlanta: $('#codigoPlanta').val(),
                    codigoCliente: $('#codigoCliente').val(),
                    codigoLoja: $('#codigoLoja').val(),
                    tipoFrete: $('#tipoFrete').val(),
                    condicaoPagamento: $('#condicaoPagamento').val(),
                    codigoProduto: codigoProduto
                },
                dataType: 'json',
                success: function(response) {
                    if (!response.sucesso) {
                        console.error('Erro ao obter valor unitário do produto: ' + response.mensagem);
                        return;
                    }

                    const valorUnitario = response.valor || 0;
                    if (valorUnitario <= 0) return;

                    ConfirmationModal.warning({
                        message: `Existe orçamento aprovado para este item com o valor de ${formatarMoeda(valorUnitario)}. Deseja alterar para este valor?`,
                        type: 'warning',
                        confirmText: 'Alterar',
                        cancelText: 'Manter valor atual',
                        onConfirm: () => {
                            $('#itemValorUnitario').maskMoney('mask', Number(String(valorUnitario).replace(',', '.')));
                            calcularTotaisItem();
                        },
                        onCancel: () => {
                        }
                    });
                },
                error: function(xhr, status, error) {
                    console.error('Erro ao obter valor unitário do produto:', error);
                }
            });
        }

        function preencherDadosProduto(dados) {
            if (dados.tipo) $('#itemTipo').val(dados.tipo).trigger('change.select2');

            $('#lotePiloto').prop('checked', dados.lotePiloto);
            $('#itemCodigoProduto').val(dados.codigo);
            $('#itemReferencia').val(dados.referencia);
            $('#itemEstrutura').val(dados.estrutura);
            $('#itemLargura').val(dados.largura);
            $('#itemAltura').val(dados.altura);
            $('#itemComprimento').val(dados.comprimento);
            $('#itemValorUnitario').maskMoney('mask', Number(String(dados.valorUnitario).replace(',', '.')));
            $('#itemVinco1').val(dados.vinco1);
            $('#itemVinco2').val(dados.vinco2);
            $('#itemVinco3').val(dados.vinco3);
            $('#itemVinco4').val(dados.vinco4);
            $('#itemVinco5').val(dados.vinco5);
            $('#itemGramatura').val(dados.gramatura || dados.gramaturaBobina).trigger('change.select2');
            $('#itemDiametroBobina').val(dados.diametroBobina);
            $('#itemTubete').val(dados.tubeteBobina).trigger('change.select2');
            $('#itemLarguraBobina').val(dados.larguraBobina);

            // Campos ocultos
            $('#itemAreaRefile').val(dados.areaRefile);
            $('#itemPeso').val(dados.peso || 0);
            $('#itemArea').val(dados.area || 0);
            $('#itemProdutoEstoque').val(dados.estoque || 0);
            $('#itemProdutoTipo').val(dados.tipoProduto || '');
            ObterValorUnitarioProdutoOrcamento();
        }

        function abrirModalRoteiroProduto(index) {
            const item = index !== undefined ? listaItens[index]: { roteiroProduto: roteiroProdutoSelecionado };
            if (!item || !$('#itemCodigoProduto').val() && index === undefined) {
                Toast.warning('Selecione um produto para visualizar o roteiro.');
                return;
            }

            const $conteudo = $('#roteiroProdutoConteudo');
            const roteiros = item.roteiroProduto || [];

            if (roteiros.length === 0) {
                $conteudo.html(`
                    <div class="text-center text-muted py-4">
                        Nenhum roteiro encontrado para este produto.
                    </div>
                `);
            } else {
                let html = '';
                roteiros.forEach(function(roteiro) {
                    const maquinas = roteiro.maquinas || [];
                    html += `
                        <div class="card mb-3">
                            <div class="card-header">
                                <div class="fw-bold">${(roteiro.referenciaItem || '-')}</div>
                                <small class="text-muted">Código: ${(roteiro.codigoItem || '-')}</small>
                            </div>
                            <div class="table-responsive">
                                <table class="table table-sm table-hover mb-0">
                                    <thead class="table-light">
                                        <tr>
                                            <th style="width: 100px;">Etapa</th>
                                            <th>Código</th>
                                            <th>Máquina</th>
                                            <th>Status</th>
                                        </tr>
                                    </thead>
                                    <tbody>
                    `;
                    if (maquinas.length === 0) {
                        html += `
                            <tr>
                                <td colspan="4" class="text-center text-muted py-3">Nenhuma máquina encontrada</td>
                            </tr>
                        `;
                    } else {
                        maquinas.forEach(function(maquina, maquinaIndex) {
                            html += `
                                <tr>
                                    <td>${maquinaIndex + 1}</td>
                                    <td>${(maquina.codigo || '-')}</td>
                                    <td>${(maquina.nome || '-')}</td>
                                    <td>${(maquina.status || '-')}</td>
                                </tr>
                            `;
                        });
                    }
                    html += `
                                    </tbody>
                                </table>
                            </div>
                        </div>
                    `;
                });
                $conteudo.html(html);
            }

            const modal = new bootstrap.Modal(document.getElementById('modalRoteiroProduto'));
            modal.show();
        }

        function ajustarSobreposicaoModalRoteiro() {
            $('#modalRoteiroProduto').on('shown.bs.modal', function() {
                $('.modal-backdrop').last().addClass('roteiro-produto-backdrop');
            });
            $('#modalRoteiroProduto').on('hidden.bs.modal', function() {
                $('.modal-backdrop').removeClass('roteiro-produto-backdrop');
                if ($('#modalItem').hasClass('show')) {
                    $('body').addClass('modal-open');
                }
            });
        }

        function carregarGramaturas() {
            const codigoPlanta = $('#codigoPlanta').val();
            const $select = $('#itemGramatura');

            $select.empty().append('<option value="">Selecione...</option>');

            if (!codigoPlanta) return;

            return $.ajax({
                url: 'EmissaoPedidoRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: { acao: 'listarGramaturas', codigoPlanta: codigoPlanta },
                dataType: 'json',
                success: function(response) {
                    if (!response.sucesso) {
                        console.error('Erro ao carregar gramaturas: ' + response.mensagem);
                        return;
                    }

                    (response.dados || []).forEach(function(item) {
                        $select.append($('<option>', {
                            value: item.codigo,
                            text: item.descricao
                        }));
                    });
                },
                error: function(xhr, status, error) {
                    console.error('Erro ao carregar gramaturas:', error);
                }
            });
        }

        function carregarCondicoesPagamento(codigoPlanta, condicaoPagamento) {
            const $select = $('#condicaoPagamento');
            $select.empty().append('<option value="">Selecione a condição de pagamento</option>');

            if (!codigoPlanta || !condicaoPagamento) return;

            return $.ajax({
                url: 'EmissaoPedidoRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: { 
                    acao: 'listarCondicoesPagamento',
                    codigoPlanta: codigoPlanta,
                    condicaoPagamento: condicaoPagamento
                },
                dataType: 'json',
                success: function(response) {
                    if (!response.sucesso) {
                        console.error('Erro ao carregar condições de pagamento: ' + response.mensagem);
                        return;
                    }

                    (response.dados || []).forEach(function(item) {
                        $select.append($('<option>', {
                            value: item.codigo,
                            text: item.descricao,
                            selected: true
                        }));
                    });
                },
                error: function(xhr, status, error) {
                    console.error('Erro ao carregar condições de pagamento:', error);
                }
            });
        }

        function carregarTubetes(codigoPlanta) {
            const $select = $('#itemTubete');
            $select.empty().append('<option value="">Selecione o tubete</option>');

            if (!codigoPlanta) return;

            return $.ajax({
                url: 'EmissaoPedidoRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: { acao: 'listarTubetes', codigoPlanta: codigoPlanta },
                dataType: 'json',
                success: function(response) {
                    if (!response.sucesso) {
                        console.error('Erro ao carregar tubetes: ' + response.mensagem);
                        return;
                    }

                    (response.dados || []).forEach(function(item) {
                        $select.append($('<option>', {
                            value: item.codigo,
                            text: item.descricao
                        }));
                    });
                },
                error: function(xhr, status, error) {
                    console.error('Erro ao carregar tubetes:', error);
                }
            });
        }

        function carregarCombosPorPlanta() {
            const codigoPlanta = $('#codigoPlanta').val();
            //carregarParametros(codigoPlanta);
            carregarGramaturas();
            carregarTubetes(codigoPlanta);
            carregarTiposItem(codigoPlanta)
        }

        function irProximo() {
            navegarParaAba(abaAtual + 1);
        }

        function irAnterior() {
            navegarParaAba(abaAtual - 1);
        }

        function anexarEventos() {
            $('#btnSalvar').on('click', function() {
                validarFormulario();
            });

            $('#formularioItem').on('submit', function(e) {
                e.preventDefault();
                $('#btnSalvarItem').trigger('click');
            });

            $('.tab-button').on('click', function() {
                const tabNumber = parseInt($(this).data('tab'));
                if (tabNumber === abaAtual) return;
                navegarParaAba(tabNumber);
            });

            $('#codigoPlanta').on('change', function() {
                limparCamposAposPlanta();
                carregarCombosPorPlanta();
                exibirCampoPorPlanta();
            atualizarCamposPorTipoItem();
            });

            $('#itemQuantidade').on('input', function() {
                calcularTotaisItem();
            });
            
            $('#itemValorUnitario').on('input keyup change', function() {
                calcularTotaisItem();
            });

            $('#itemTipo').on('change', function() {
                $('#formularioItem').find('input, select, textarea').not('#itemTipo, #lotePiloto').val('');
                $('#formularioItem .select2').not('#itemTipo').val('').trigger('change.select2');
                $('#itemQuantidade, #itemValorUnitario, #itemValorTotal, #itemAreaTotal, #itemAreaRefile').val('0');
                produtoSelecionado = null;
                roteiroProdutoSelecionado = [];
                atualizarCamposPorTipoItem();
            });

            $('#btnSalvarItem').on('click', async function() {
                const validador = validacoes.modalItem;
                const formularioValido = await validador.revalidate();

                if (!formularioValido) {
                    exibirErrosValidacao(validador);
                    return;
                }

                calcularTotaisItem();

                const valorUnitario = desformatarMoeda($('#itemValorUnitario').val());

                exibirLoading(true);

                try {
                    const respostaLotePiloto = await VerificarExisteLotePiloto();
                    if (!respostaLotePiloto.sucesso) {
                        Toast.error(respostaLotePiloto.mensagem || 'Erro ao verificar lote piloto', 'Erro', 6000);
                        return;
                    } 

                    $('#lotePiloto').prop('checked', respostaLotePiloto.existe == 1);
                    if ($('#lotePiloto').is(':checked') && existeItemLotePilotoDuplicado()) {
                        Toast.warning('Só é permitido um lote piloto por produto', 'Atenção', 6000);
                        return;
                    }
                    
                    const respostaValidacao = await validarItens();
                    if (!respostaValidacao.sucesso) {
                        Toast.error( respostaValidacao.mensagem || 'Um ou mais itens inválidos', 'Erro',  6000);
                        return;
                    }

                    const respostaOrcamento = await verificarExisteOrcamento();
                    const orcamentoAprovado = respostaOrcamento.existe == 1;

                    // Validação pedido de compra
                    const pedidoCompra = String($('#itemPedidoCompra').val() || '').trim();
                    const pedidoCompraDuplicado = pedidoCompra !== '' && listaItens.some(function(item, index) {
                        return index !== itemEditandoIndex && String(item.pedidoCompra || '').trim().toUpperCase() === pedidoCompra.toUpperCase();
                    });

                    const respostaPedidoCompra = await verificarPedidoCompraDuplicado();
                    if (pedidoCompraDuplicado || respostaPedidoCompra.existe == 1) {
                        ConfirmationModal.warning({
                            message: 'Existe Pedido cadastrado com este Pedido de Compra e Item. Deseja confirmar a inclusão do item?',
                            type: 'warning',
                            confirmText: 'Confirmar',
                            cancelText: 'Cancelar',
                            onConfirm: () => {
                                salvarItem(orcamentoAprovado);
                            },
                            onCancel: () => {}
                        });
                        return;
                    }

                    const possuiMaquinaInativa = roteiroProdutoSelecionado.some(function(roteiro) {
                        return (roteiro.maquinas || []).some(function(maquina) {
                            return String(maquina.status || '').trim().toUpperCase() === 'INATIVA';
                        });
                    });
                    
                    if (possuiMaquinaInativa) {
                        ConfirmationModal.warning({
                            message: 'Existe uma máquina inativa no roteiro do item. Deseja confirmar a inclusão do item?',
                            type: 'warning',
                            confirmText: 'Confirmar',
                            cancelText: 'Cancelar',
                            onConfirm: () => {
                                salvarItem(orcamentoAprovado);
                            },
                            onCancel: () => {}
                        });
                        return;
                    }
                    
                    salvarItem(orcamentoAprovado);
                } catch (erro) {
                    console.error('Erro ao validar os itens:', erro);
                    Toast.error('Não foi possível validar os itens.');
                } finally {
                    exibirLoading(false);
                }
            });

            let buscaClienteTimeout = null;
            $('#clienteBusca').on('input', function() {
                clearTimeout(buscaClienteTimeout);
                buscaClienteTimeout = setTimeout(function() { pesquisarCliente(); }, 500);
            });

            $('#clienteBusca').on('keypress', function(e) {
                if (e.which === 13) {
                    e.preventDefault();
                    clearTimeout(buscaClienteTimeout);
                    pesquisarCliente();
                }
            });

            let buscaClienteRemessaTimeout = null;
            $('#clienteRemessaBusca').on('input', function() {
                clearTimeout(buscaClienteRemessaTimeout);
                buscaClienteRemessaTimeout = setTimeout(function() { pesquisarClienteRemessa(); }, 500);
            });

            $('#clienteRemessaBusca').on('keypress', function(e) {
                if (e.which === 13) {
                    e.preventDefault();
                    clearTimeout(buscaClienteRemessaTimeout);
                    pesquisarClienteRemessa();
                }
            });

            let buscaProdutoTimeout = null;
            $('#produtoBusca').on('input', function() {
                clearTimeout(buscaProdutoTimeout);
                buscaProdutoTimeout = setTimeout(function() { pesquisarProduto(); }, 500);
            });

            $('#produtoBusca').on('keypress', function(e) {
                if (e.which === 13) {
                    e.preventDefault();
                    clearTimeout(buscaProdutoTimeout);
                    pesquisarProduto();
                }
            });
        }
        
        function gerarIdTemporario() {
            if (window.crypto && typeof window.crypto.randomUUID === 'function') {
                return window.crypto.randomUUID();
            }

            return Date.now().toString(36) + '-' + Math.random().toString(36).substring(2, 12);
        }

        function ocultarFeedbackInsercao() {
            $('#feedbackInsercaoPedidos')
                .addClass('d-none')
                .removeClass('alert-warning alert-danger alert-info')
                .empty();
        }

        function descricaoItem(item) {
            if (!item) return 'Item nao identificado';

            const partes = [];
            if (item.codigoProduto) partes.push(`Produto ${item.codigoProduto}`);
            if (item.referencia) partes.push(`Referencia ${item.referencia}`);

            return partes.length > 0 ? partes.join(' - ') : 'Item nao identificado';
        }

        function mostrarFeedbackInsercao(response, itensAntesDoEnvio) {
            const resultados = Array.isArray(response.resultados) ? response.resultados : [];
            const itensPorId = new Map(
                (itensAntesDoEnvio || []).map(item => [String(item.idTemporario || ''), item])
            );
            const gravados = resultados.filter(resultado => resultado.gravado === true);
            const naoGravados = resultados.filter(resultado => resultado.gravado !== true);
            const classe = gravados.length > 0 ? 'alert-warning' : 'alert-danger';

            const resumo = gravados.length > 0
                ? `${gravados.length} pedido(s) foram criado(s) com sucesso. ${naoGravados.length} item(ns) nao puderam ser processado(s) e continuam disponiveis para correcao.`
                : 'Nenhum pedido foi criado. Verifique as ocorrencias abaixo, corrija os dados e tente novamente.';

            const codigosGeradosHtml = gravados.length > 0 ? `
                <div class="mt-3">
                    <strong>Pedidos criados com sucesso:</strong>
                    <span>${gravados.map(resultado => (resultado.pedido)).join(', ')}</span>
                </div>
            ` : '';

            const errosHtml = naoGravados.length > 0 ? `
                <div class="mt-3">
                    <strong>Itens que precisam de atencao:</strong>
                    <ul class="mb-0 mt-1">
                        ${naoGravados.map(resultado => {
                            const item = itensPorId.get(String(resultado.idTemporario || ''));
                            return `
                                <li>
                                    ${(resultado.mensagem || 'Nao foi possivel processar este item.')}
                                </li>
                            `;
                        }).join('')}
                    </ul>
                </div>
            ` : '';

            $('#feedbackInsercaoPedidos')
                .removeClass('d-none alert-warning alert-danger alert-info')
                .addClass(classe)
                .html(`
                    <div class="d-flex align-items-start gap-2">
                        <i class="bi bi-exclamation-triangle-fill mt-1"></i>
                        <div class="flex-grow-1">
                            <h6 class="alert-heading mb-2">Resultado do processamento</h6>
                            <p class="mb-0">${(resumo)}</p>
                            ${codigosGeradosHtml}
                            ${errosHtml}
                            ${naoGravados.length > 0 ? '<p class="mb-0 mt-3">Revise os itens indicados e envie o formulario novamente.</p>' : ''}
                        </div>
                    </div>
                `);

            document.getElementById('feedbackInsercaoPedidos').scrollIntoView({
                behavior: 'smooth',
                block: 'start'
            });
        }

        function mostrarFalhaComunicacao(mensagem) {
            $('#feedbackInsercaoPedidos')
                .removeClass('d-none alert-warning alert-info')
                .addClass('alert-danger')
                .html(`
                    <div class="d-flex align-items-start gap-2">
                        <i class="bi bi-x-circle-fill mt-1"></i>
                        <div>
                            <h6 class="alert-heading mb-2">Nao foi possivel confirmar o processamento</h6>
                            <p class="mb-0">${(mensagem)}</p>
                            <p class="mb-0 mt-2"><strong>Antes de reenviar, consulte a grade para evitar pedidos duplicados.</strong></p>
                        </div>
                    </div>
                `);
        }

        function removerItensGravados(resultados) {
            const idsGravados = new Set(
                (resultados || [])
                    .filter(resultado => resultado.gravado === true)
                    .map(resultado => String(resultado.idTemporario || ''))
            );

            if (idsGravados.size === 0) return;

            listaItens = listaItens.filter(item => !idsGravados.has(String(item.idTemporario || '')));
            renderizarListaItens();
        }

        async function validarFormulario() {
            const abas = [1, 2];
            const resultados = await Promise.all([
                validacoes.aba1.revalidate(),
                Promise.resolve(listaItens.length > 0)
            ]);

            resultados.forEach((valida, index) => {
                const numeroAba = abas[index];
                if (valida) {
                    marcarAbaComoCompleta(numeroAba);
                } else {
                    marcarAbaComoInvalida(numeroAba);
                    if (validacoes[`aba${numeroAba}`]) {
                        exibirErrosValidacao(validacoes[`aba${numeroAba}`]);
                    }
                }
            });

            if (!resultados[1]) {
                Toast.error('Adicione ao menos um item.');
            }

            if (!resultados.every(Boolean)) {
                const primeiraInvalida = abas[resultados.findIndex(v => !v)];
                navegarParaAba(primeiraInvalida);
                Toast.error('Existem campos obrigatórios pendentes. Verifique as seções destacadas.');
                return;
            }

            enviarFormulario();
        }

        function enviarFormulario() {
            const dadosFormulario = {
                acao: id ? 'editar' : 'criar',
                codigoPlanta: $('#codigoPlanta').val(),
                codigoCliente: $('#codigoCliente').val(),
                codigoLoja: $('#codigoLoja').val(),
                razaoSocial: $('#razaoSocial').val(),
                tes: $('#tes').val(),
                condicaoPagamento: $('#condicaoPagamento').val(),
                moeda: $('#moeda').val(),
                tipoFrete: $('#tipoFrete').val(),
                codigoClienteRemessa: $('#codigoClienteRemessa').val(),
                codigoLojaRemessa: $('#codigoLojaRemessa').val(),
                validaLimiteCredito: $('#hidValidaLimiteCredito').val(),
                toleranciaMaior: $('#hidToleranciaMaior').val(),
                toleranciaDataMaior: $('#hidToleranciaDataMaior').val(),
                tesRemessa: $('#hidTesRemessa').val(),
                itens: JSON.stringify(
                    [...listaItens].map(({ roteiroProduto, ...item }) => item)
                )
            };

            if (id) {
                dadosFormulario.id = id;
            }

            const itensAntesDoEnvio = listaItens.map(item => ({ ...item }));

            ocultarFeedbackInsercao();
            exibirLoading(true);
            $('#btnSalvar').prop('disabled', true);

            $.ajax({
                url: 'EmissaoPedidoRepositorio.asp',
                type: 'POST',
                data: dadosFormulario,
                dataType: 'json',
                success: function(response) {
                    const resultados = Array.isArray(response.resultados) ? response.resultados : [];

                    if (response.sucesso === true) {
                        Toast.success(response.mensagem || 'Pedidos inseridos com sucesso!');
                        setTimeout(function() { navegarParaGrade(); }, 500);
                        return;
                    }

                    removerItensGravados(resultados);
                    mostrarFeedbackInsercao(response, itensAntesDoEnvio);
                    navegarParaAba(2);
                },
                error: function(xhr, status, error) {
                    console.error('Erro ao salvar formulario:', status, error, xhr.responseText);
                    mostrarFalhaComunicacao(
                        'O servidor nao devolveu uma resposta valida. Os itens foram mantidos na tela, pois nao foi possivel identificar quais pedidos foram gravados.'
                    );
                },
                complete: function() {
                    exibirLoading(false);
                    $('#btnSalvar').prop('disabled', false);
                }
            });
        }

        function limparCamposItem() {
            $('#formularioItem')[0].reset();
            $('#formularioItem .select2').val('').trigger('change.select2');
            $('#itemTipo').val('').trigger('change.select2');
            $('#itemValorTotal').val('');
            $('#lotePiloto').prop('checked', false).prop('disabled', true);
            roteiroProdutoSelecionado = [];

            exibirCampoPorPlanta();

            if (validacoes.modalItem) validacoes.modalItem.refresh();
        }

        function abrirModalItem() {
            itemEditandoIndex = null;
            limparCamposItem();

            $('#modalItemTitulo').text('Adicionar Item');
            $('#btnSalvarItem').html('<i class="bi bi-plus-lg me-1"></i>Adicionar');
            const modal = new bootstrap.Modal(document.getElementById('modalItem'));
            modal.show();
        }

        function editarItem(index) {
            const item = listaItens[index];
            if (!item) return;
            
            itemEditandoIndex = index;
            limparCamposItem();

            const modal = new bootstrap.Modal(document.getElementById('modalItem'));
            modal.show();

            $('#lotePiloto').prop('checked', item.lotePiloto == 1 || item.lotePiloto === true);
            $('#itemTipo').val(item.itemTipo || '').trigger('change.select2');
            $('#itemCodigoProduto').val(item.codigoProduto || '');
            $('#itemReferencia').val(item.referencia || '');
            $('#itemEstrutura').val(item.estrutura || '');
            $('#itemQuantidade').val(item.quantidade || '');
            $('#itemValorUnitario').maskMoney('mask', Number(String(item.valorUnitario).replace(',', '.')));
            $('#itemValorTotal').val(item.valorTotal || 0);
            $('#itemAreaTotal').val(item.areaTotal || 0);
            $('#itemGramatura').val(item.gramatura || 0).trigger('change.select2');
            $('#itemLargura').val(item.largura || 0);
            $('#itemComprimento').val(item.comprimento || 0);
            $('#itemVinco1').val(item.vinco1 || 0);
            $('#itemVinco2').val(item.vinco2 || 0);
            $('#itemVinco3').val(item.vinco3 || 0);
            $('#itemVinco4').val(item.vinco4 || 0);
            $('#itemVinco5').val(item.vinco5 || 0);
            $('#itemDiametroBobina').val(item.diametroBobina || 0);
            $('#itemLarguraBobina').val(item.larguraBobina || 0);
            $('#itemTubete').val(item.tubeteBobina || '').trigger('change.select2');
            $('#itemDataEntrega').val(item.dataEntrega || '');
            $('#itemDataPrevistaEntrega').val(item.dataPrevistaEntrega || '');
            $('#itemModalFrete').val(item.modalFrete || '').trigger('change.select2');
            $('#itemComercioExterior').val(item.comercioExterior || '').trigger('change.select2');
            $('#itemPedidoCompra').val(item.pedidoCompra || '');
            $('#itemPedido').val(item.itemPedido || '');
            $('#itemObservacao').val(item.observacao || '');
            $('#itemAreaRefile').val(item.areaRefile || '');
            $('#itemArea').val(item.area || 0);
            $('#itemPeso').val(item.peso || 0);
            $('#itemProdutoTipo').val(item.tipoProduto || '');
            roteiroProdutoSelecionado = item.roteiroProduto || [];
            atualizarCamposPorTipoItem();

            $('#modalItemTitulo').text('Editar Item');
            $('#btnSalvarItem').html('<i class="bi bi-check-lg me-1"></i>Salvar');
        }

        function removerItem(index) {
            ConfirmationModal.danger({
                message: 'Deseja realmente remover este item?',
                type: 'danger',
                confirmText: 'Remover',
                cancelText: 'Cancelar',
                onConfirm: () => {
                    listaItens.splice(index, 1);
                    renderizarListaItens();
                    Toast.success('Item removido com sucesso!');
                },
                onCancel: () => {}
            });
        }

        function salvarItem(orcamentoAprovado) {
            calcularTotaisItem();

            const itemAnterior = itemEditandoIndex !== null ? listaItens[itemEditandoIndex] : null;

            const item = {
                idTemporario: itemAnterior && itemAnterior.idTemporario ? itemAnterior.idTemporario : gerarIdTemporario(),
                lotePiloto: $('#lotePiloto').is(':checked') ? 1 : 0,
                orcamentoAprovado: orcamentoAprovado,
                itemTipo: $('#itemTipo').val(),
                codigoProduto: $('#itemCodigoProduto').val(),
                tipoProduto: $('#itemProdutoTipo').val(),
                referencia: $('#itemReferencia').val(),
                estrutura: $('#itemEstrutura').val(),
                quantidade: desformatarMoeda($('#itemQuantidade').val()),
                valorUnitario: desformatarMoeda($('#itemValorUnitario').val()),
                valorTotal: desformatarMoeda($('#itemValorTotal').val()),
                areaTotal: desformatarMoeda($('#itemAreaTotal').val()),
                areaRefile: $('#itemAreaRefile').val(),
                area: $('#itemArea').val(),
                peso: $('#itemPeso').val(),
                gramatura: $('#itemGramatura').val(),
                largura: $('#itemLargura').val(),
                comprimento: $('#itemComprimento').val(),
                vinco1: $('#itemVinco1').val(),
                vinco2: $('#itemVinco2').val(),
                vinco3: $('#itemVinco3').val(),
                vinco4: $('#itemVinco4').val(),
                vinco5: $('#itemVinco5').val(),
                diametroBobina: $('#itemDiametroBobina').val(),
                larguraBobina: $('#itemLarguraBobina').val(),
                tubeteBobina: $('#itemTubete').val(),
                dataEntrega: $('#itemDataEntrega').val(),
                dataPrevistaEntrega: $('#itemDataPrevistaEntrega').val(),
                modalFrete: $('#itemModalFrete').val() || '',
                comercioExterior: $('#itemComercioExterior').val() || '',
                pedidoCompra: $('#itemPedidoCompra').val(),
                itemPedido: $('#itemPedido').val(),
                observacao: $('#itemObservacao').val(),
                roteiroProduto: roteiroProdutoSelecionado,
            };

            if (itemEditandoIndex !== null) {
                listaItens[itemEditandoIndex] = item;
                Toast.success('Item atualizado com sucesso!');
            } else {
                listaItens.push(item);
                Toast.success('Item adicionado com sucesso!');
            }

            renderizarListaItens();
            const modal = bootstrap.Modal.getInstance(document.getElementById('modalItem'));
            modal.hide();
        }

        function renderizarListaItens() {
            const $container = $('#itensListaContainer');
            const $contador = $('#contadorItens, #menuContadorItens');
            $contador.text(listaItens.length);

            if (listaItens.length === 0) {
                $container.html(`
                    <div class="text-center py-5">
                        <p class="text-muted mt-3 mb-0">Nenhum item adicionado.</p>
                    </div>
                `);
                return;
            }

            const totais = listaItens.reduce((acc, item) => {
                acc.quantidade += desformatarMoeda(item.quantidade);
                acc.area += desformatarMoeda(item.areaTotal);
                acc.valorTotal += desformatarMoeda(item.valorTotal);
                return acc;
            }, { quantidade: 0, area: 0, valorUnitario: 0, valorTotal: 0 });

            let html = `
                <div class="table-responsive">
                    <table id="gradeItem" class="table table-hover table-sm mb-0">
                        <thead class="table-light">
                            <tr>
                                <th class="text-center">Lote Piloto</th>
                                <th class="text-center">Orçamento</th>
                                <th>Código Produto</th>
                                <th>Referência</th>
                                <th class="text-end">Quantidade</th>
                                <th class="text-end">Área</th>
                                <th class="text-end">Valor Unitário</th>
                                <th class="text-end">Valor Total</th>
                                <th>Pedido de Compra</th>
                                <th>Item Pedido</th>
                                <th>Data de Entrega</th>
                                <th class="text-center" style="width: 100px;">Ações</th>
                            </tr>
                        </thead>
                        <tbody>
            `;

            listaItens.forEach((item, index) => {
                html += `
                    <tr>
                        <td class="text-center">
                            <span class="badge ${item.lotePiloto == '1' ? 'bg-success' : 'bg-secondary'}">
                                ${item.lotePiloto == 1 ? 'Sim' : 'Não'}
                            </span>
                        </td>
                        <td class="text-center">
                            ${item.orcamentoAprovado
                                ? '<span class="badge bg-success">Aprovado</span>'
                                : '<span class="badge bg-warning text-dark" title="Item sem orçamento aprovado"><i class="bi bi-exclamation-triangle-fill me-1"></i>Atenção</span>'}
                        </td>
                        <td>${item.codigoProduto}</td>
                        <td>${item.referencia}</td>
                        <td class="text-end">${formatarMoeda(item.quantidade, 0)}</td>
                        <td class="text-end">${formatarMoeda(item.areaTotal)}</td>
                        <td class="text-end">
                         ${item.valorUnitario > 50
                                ? '<i class="bi bi-exclamation-circle-fill text-dark ms-1" title="Valor unitário parece alto"></i>'
                                : ''}
                            ${formatarMoeda(item.valorUnitario)} 
                        </td>
                        <td class="text-end">${formatarMoeda(item.valorTotal)}</td>
                        <td>${item.pedidoCompra || '-'}</td>
                        <td>${item.itemPedido || '-'}</td>
                        <td>${formatarParaDataBR(`${item.dataEntrega} 00:00:00`, false)}</td>
                        <td style="text-align: center">
                            <dropdown-menu label="Ações">
                                <a class="dropdown-item" onclick="abrirModalRoteiroProduto(${index})"><i class="bi bi-diagram-3 me-1"></i> Roteiro de Produto</a>
                                <a class="dropdown-item" onclick="editarItem(${index})"><i class="bi bi-pencil me-1"></i> Editar</a>
                                <a class="dropdown-item" onclick="removerItem(${index})"><i class="bi bi-trash me-1"></i> Excluir</a>
                            </dropdow-menu>
                        </td>
                    </tr>
                `;
            });

            html += `
                    </tbody>
                        <tfoot class="table-light fw-bold">
                            <tr>
                                <td colspan="4"">Totais</td>
                                <td class="text-end">${formatarMoeda(totais.quantidade, 0)}</td>
                                <td class="text-end">${formatarMoeda(totais.area)}</td>
                                <td class="text-end"></td>
                                <td class="text-end">${formatarMoeda(totais.valorTotal)}</td>
                                <td></td>
                                <td></td>
                                <td></td>
                                <td></td>
                            </tr>
                        </tfoot>
                    </table>
                </div>
            `;
            $container.html(html);
        }


        $(document).ready(function() {
            $('#titulo').text('Emissão de Pedidos');

            anexarEventos();
            navegarParaAba(1);
            inicializarSelect2();
            inicializarMascaras();
            inicializarValidacoes();
            ajustarSobreposicaoModalRoteiro();
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

            <div id="divAcoes" class="d-flex mb-0 page-actions flex-wrap">
                <button type="button" class="btn btn-sm btn-success" id="btnSalvar">
                    <i class="bi bi-check-lg me-1"></i>
                    Enviar Formulário
                </button>
            </div>
        </div>

        <div id="loadingOverlay" class="position-fixed top-0 start-0 w-100 h-100 d-none justify-content-center align-items-center" style="background: rgba(255,255,255,0.8); z-index: 9999;">
            <div class="spinner-border text-primary" role="status">
                <span class="visually-hidden">Carregando...</span>
            </div>
        </div>

        <!-- Feedback permanente do processamento parcial. -->
        <div id="feedbackInsercaoPedidos" class="alert d-none" role="alert" aria-live="assertive"></div>

        <div class="form-container">
            <div class="form-layout">
                <div class="tabs-container">
                    <div class="tabs">
                        <div class="tab-group-title">Geral</div>

                        <button type="button" class="tab-button active" data-tab="1">
                            <i class="tab-icon bi bi-clipboard-check"></i>
                            <span class="tab-title">Dados Gerais</span>
                        </button>

                        <button type="button" class="tab-button" data-tab="2">
                            <i class="tab-icon bi bi-receipt"></i>
                            <span class="tab-title">Itens</span>
                            <span class="tab-meta" id="menuContadorItens">0</span>
                        </button>
                    </div>
                </div>

                <div class="tab-content-wrapper">
                    <form id="formularioPedido" name="formularioPedido" novalidate>
                        <!-- Campos Hidden TES -->
                        <input type="hidden" name="hidTesRemessa" id="hidTesRemessa" value="N"/>
                        <input type="hidden" name="hidPoder3" id="hidPoder3" value="N"/>
                        <input type="hidden" name="hidValidaLimiteCredito" id="hidValidaLimiteCredito" value="N"/>

                        <!-- Campos Cliente -->
                        <input type="hidden" name="hidToleranciaMaior" id="hidToleranciaMaior" value="0"/>
                        <input type="hidden" name="hidToleranciaDataMaior" id="hidToleranciaDataMaior" value="0"/>

                        <div class="tab-content active" data-tab-content="1">
                            <div class="section-header">
                                <div class="section-title">Dados Gerais</div>
                                <p class="section-subtitle">Dados gerais do pedido.</p>
                            </div>

                            <div class="row g-3">
                                 <div class="col-12">
                                    <label for="codigoPlanta" class="form-label">Planta <span class="required-field">*</span></label>
                                    <select class="form-select form-select-sm select2" name="codigoPlanta" id="codigoPlanta">
                                        <option value="">Selecione um item</option>
                                        <option value="05">BENTO GONÇALVES</option>
                                        <option value="06">MARANGUAPE</option>
                                        <option value="07">SAQUAREMA</option>
                                    </select>
                                </div>

                                <div class="col-12 col-md-4">
                                    <label for="codigoCliente" class="form-label">Cód. Cliente <span class="required-field me-2">*</span></label>
                                    <div class="input-group">
                                        <button type="button" class="btn btn-sm btn-secondary" id="btnBuscarCliente" onclick="abrirModalCliente()">
                                            Buscar
                                            <i class="bi bi-box-arrow-up-right"></i>
                                        </button>
                                        <input disabled type="text" class="form-control form-control-sm" id="codigoCliente" name="codigoCliente">
                                        <button type="button" class="btn btn-sm btn-outline-secondary" id="btnLimiteCliente" onclick="abrirModalLimiteCliente()" title="Visualizar limite de crédito">
                                            <i class="bi bi-cash-coin"></i> Limite
                                        </button>
                                    </div>
                                    <small class="text-muted">Clique no botão para buscar.</small>
                                </div>
                                
                                <div class="col-12 col-md-4">
                                    <label for="codigoLoja" class="form-label">Cód. Loja <span class="required-field">*</span></label>
                                    <input disabled type="text" class="form-control form-control-sm" id="codigoLoja" name="codigoLoja">
                                </div>
                                
                                <div class="col-12 col-md-4">
                                    <label for="razaoSocial" class="form-label">Nome Cliente <span class="required-field">*</span></label>
                                    <input disabled type="text" class="form-control form-control-sm" name="razaoSocial" id="razaoSocial" maxlength="160" />
                                </div>

                                <div class="col-12 col-md-4">
                                    <label for="tes" class="form-label">TES <span class="required-field me-2">*</span></label>
                                    <div class="input-group">
                                        <button type="button" class="btn btn-sm btn-secondary" id="btnBuscarTes" onclick="abrirModalTes()">
                                            Buscar
                                            <i class="bi bi-box-arrow-up-right"></i>
                                        </button>
                                        <input disabled type="text" class="form-control form-control-sm" id="tes" name="tes">
                                    </div>
                                    <small class="text-muted">Clique no botão para buscar.</small>
                                </div>

                                <div class="col-12 col-md-4 campos-remessa d-none">
                                    <label for="codigoClienteRemessa" class="form-label">Cód. Cliente Remessa <span class="required-field me-2">*</span></label>
                                    <div class="input-group">
                                        <button type="button" class="btn btn-sm btn-secondary" id="btnBuscarClienteRemessa" onclick="abrirModalClienteRemessa()">
                                            Buscar
                                            <i class="bi bi-box-arrow-up-right"></i>
                                        </button>
                                        <input disabled type="text" class="form-control form-control-sm" id="codigoClienteRemessa" name="codigoClienteRemessa">
                                    </div>
                                    <small class="text-muted">Clique no botão para buscar.</small>
                                </div>

                                <div class="col-12 col-md-4 campos-remessa d-none">
                                    <label for="codigoLojaRemessa" class="form-label">Cód. Loja Remessa <span class="required-field me-2">*</span></label>
                                    <input type="text" class="form-control form-control-sm" id="codigoLojaRemessa" name="codigoLojaRemessa" readonly>
                                </div>

                                <div class="col-12 col-md-4 campos-remessa d-none">
                                    <label for="nomeClienteRemessa" class="form-label">Nome Cliente Remessa <span class="required-field me-2">*</span></label>
                                    <input type="text" class="form-control form-control-sm" id="nomeClienteRemessa" name="nomeClienteRemessa" readonly>
                                </div>

                                <div class="col-12 col-md-4 campo-saldo-aparas d-none">
                                    <label for="saldoAparas" class="form-label">Saldo Aparas <span class="required-field me-2">*</span></label>
                                    <input type="text" class="form-control form-control-sm" id="saldoAparas" name="saldoAparas" value="0" readonly>
                                </div>  
                
                                <div class="col-12 col-md-4">
                                    <label for="condicaoPagamento" class="form-label">Condição de Pagamento <span class="required-field">*</span></label>
                                    <select class="form-select form-select-sm select2" name="condicaoPagamento" id="condicaoPagamento">
                                        <option value="">Selecione a condição de pagamento</option>
                                    </select>
                                </div>

                                <div class="col-12 col-md-4">
                                    <label for="moeda" class="form-label">Moeda <span class="required-field">*</span></label>
                                    <select class="form-select form-select-sm select2" name="moeda" id="moeda">
                                        <option value="">Selecione a moeda</option>
                                        <option selected value="0">REAL</option>
                                        <option value="4">DOLAR</option>
                                    </select>
                                </div>

                                <div class="col-12 col-md-4">
                                    <label for="tipoFrete" class="form-label">Tipo de Frete <span class="required-field">*</span></label>
                                    <select class="form-select form-select-sm select2" name="tipoFrete" id="tipoFrete">
                                        <option value="">Selecione o frete</option>
                                        <option value="0">CIF</option>
                                        <option value="1">FOB</option>
                                    </select>
                                </div>

                            </div>
                        </div>
                        
                        <div class="tab-content" data-tab-content="2">
                            <div class="section-header d-flex justify-content-between align-items-center">
                                <div>
                                    <div class="section-title">
                                        Itens do Pedido <span class="badge bg-secondary ms-1" id="contadorItens">0</span>
                                    </div>
                                    <p class="section-subtitle">Itens Vinculados ao Pedido.</p>
                                </div>

                                <button type="button" class="btn btn-sm btn-primary" id="btnAdicionarNota" onclick="abrirModalItem()">
                                    <i class="bi bi-plus-lg me-1"></i>Adicionar
                                </button>
                            </div>

                            <div id="itensListaContainer"></div>
                        </div>
                    </form>

                    <div class="nav-footer">
                        <button type="button" id="btnVoltarPagina" class="btn btn-sm btn-outline-secondary" onclick="irAnterior()">
                            <i class="bi bi-arrow-left"></i> Voltar
                        </button>
                        <button type="button" id="btnProximaPagina" class="btn btn-sm btn-primary" onclick="irProximo()">
                            Próximo <i class="bi bi-arrow-right"></i>
                        </button>
                    </div>
                </div>
            </div>
        </div>
    </div>
    
    <!-- Modal de Roteiro do Produto -->
    <div class="modal fade" id="modalRoteiroProduto" tabindex="-1" aria-labelledby="modalRoteiroProdutoLabel" aria-hidden="true">
        <div class="modal-dialog modal-xl modal-dialog-centered modal-dialog-scrollable">
            <div class="modal-content">
                <div class="modal-header">
                    <h5 class="modal-title" id="modalRoteiroProdutoLabel">
                        <i class="bi bi-diagram-3 me-2"></i>Roteiro do Produto
                    </h5>
                    <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Fechar"></button>
                </div>
                <div class="modal-body">
                    <div id="roteiroProdutoConteudo"></div>
                </div>
            </div>
        </div>
    </div>

    <!-- Modal de Item -->
    <div class="modal fade" id="modalItem" tabindex="-1" aria-labelledby="modalItemLabel" aria-hidden="true">
        <div class="modal-dialog modal-xl">
            <div class="modal-content">
                <div class="modal-header">
                    <h5 class="modal-title" id="modalItemTitulo">
                        <i class="bi bi-puzzle me-2"></i>Adicionar Item
                    </h5>
                    <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Fechar"></button>
                </div>

                <div class="modal-body" style=" overflow: auto; max-height: 70vh;">
                    <form id="formularioItem" name="formularioItem" novalidate>
                        <!-- Campos Hidden produto -->
                        <input type="hidden" id="itemAreaRefile" name="itemAreaRefile" value="0"/>
                        <input type="hidden" id="itemPeso" value="0"/>
                        <input type="hidden" id="itemArea" value="0"/>
                        <input type="hidden" id="itemProdutoEstoque" value="0"/>
                        <input type="hidden" id="itemProdutoTipo" value=""/>

                        
                        <div class="switch-item pb-3 mb-4 border-bottom">
                            <div class="">
                                <label class="switch-label" for="lotePiloto">Lote Piloto</label>
                            </div>
                            <div class="form-check form-switch mb-0">
                                <input class="form-check-input" type="checkbox" id="lotePiloto" disabled/>
                            </div>
                        </div>
                            
                        <div class="row g-3">
                            <div class="col-12 col-lg-4">
                                <label for="itemTipo" class="form-label">Tipo do Item <span class="required-field">*</span></label>
                                <select class="form-select form-select-sm select2" name="itemTipo" id="itemTipo">
                                    <option value="">Selecione...</option>
                                </select>
                            </div>

                            <div class="col-12 col-lg-4">
                                <div class="d-flex align-items-center gap-1">
                                    <label for="itemCodigoProduto" class="form-label mb-0">Produto <span class="required-field">*</span></label>
                                    <button type="button"
                                        class="btn btn-link btn-sm p-0 mb-1 text-info"
                                        onclick="abrirModalRoteiroProduto()"
                                        title="Visualizar roteiro do produto"
                                        aria-label="Visualizar roteiro do produto">
                                        <i class="bi bi-info-circle-fill"></i>
                                    </button>
                                </div>
                                <div class="input-group">
                                    <button type="button" class="btn btn-sm btn-secondary" id="btnBuscarProduto" onclick="abrirModalProduto()">
                                        Buscar
                                        <i class="bi bi-box-arrow-up-right"></i>
                                    </button>
                                    <input disabled required type="text" class="form-control form-control-sm" id="itemCodigoProduto" name="itemCodigoProduto" required>
                                </div>
                                <small class="text-muted">Clique no botão para buscar.</small>
                            </div>

                            <div class="col-12 col-lg-4">
                                <label for="itemReferencia" class="form-label">Referência <span class="required-field">*</span></label>
                                <input disabled type="text" class="form-control form-control-sm" id="itemReferencia" name="itemReferencia">
                            </div>

                            <div class="col-12 col-lg-4">
                                <label for="itemEstrutura" class="form-label">Estrutura </label>
                                <input disabled type="text" class="form-control form-control-sm" id="itemEstrutura" name="itemEstrutura">
                            </div>

                            <div class="col-12 col-lg-4">
                                <label for="itemQuantidade" class="form-label">Quantidade <span class="required-field">*</span></label>
                                <input value="0" type="text" class="form-control form-control-sm" id="itemQuantidade" name="itemQuantidade">
                            </div>

                            <div class="col-12 col-lg-4">
                                <label for="itemValorUnitario" class="form-label">Valor Unitário <span class="required-field">*</span></label>
                                <input value="0" type="text" class="form-control form-control-sm" id="itemValorUnitario" name="itemValorUnitario">
                            </div>

                            <div class="col-12 col-lg-4">
                                <label for="itemValorTotal" class="form-label">Total <span class="required-field">*</span></label>
                                <input disabled value="0" type="text" class="form-control form-control-sm" id="itemValorTotal" name="itemValorTotal">
                            </div>

                            <div class="col-12 col-lg-4">
                                <label for="itemAreaTotal" class="form-label">Área Total <span class="required-field">*</span></label>
                                <input disabled value="0" type="text" class="form-control form-control-sm" id="itemAreaTotal" name="itemAreaTotal">
                            </div>

                            <div class="col-12 col-lg-4">
                                <label for="itemLargura" class="form-label">Largura</label>
                                <input value="0" type="number" class="form-control form-control-sm" id="itemLargura" name="itemLargura">
                            </div>

                            <div class="col-12 col-lg-4">
                                <label for="itemComprimento" class="form-label">Comprimento</label>
                                <input value="0" type="number" class="form-control form-control-sm" id="itemComprimento" name="itemComprimento">
                            </div>

                            <div class="col-12 col-lg-4">
                                <label for="itemVinco1" class="form-label">Vinco 1 </label>
                                <input value="0" type="number" class="form-control form-control-sm" id="itemVinco1" name="itemVinco1">
                            </div>

                            <div class="col-12 col-lg-4">
                                <label for="itemVinco2" class="form-label">Vinco 2 </label>
                                <input value="0" type="number" class="form-control form-control-sm" id="itemVinco2" name="itemVinco2">
                            </div>

                            <div class="col-12 col-lg-4">
                                <label for="itemVinco3" class="form-label">Vinco 3 </label>
                                <input value="0" type="number" class="form-control form-control-sm" id="itemVinco3" name="itemVinco3">
                            </div>

                            <div class="col-12 col-lg-4">
                                <label for="itemVinco4" class="form-label">Vinco 4 </label>
                                <input value="0" type="number" class="form-control form-control-sm" id="itemVinco4" name="itemVinco4">
                            </div>

                            <div class="col-12 col-lg-4">
                                <label for="itemVinco5" class="form-label">Vinco 5 </label>
                                <input value="0" type="number" class="form-control form-control-sm" id="itemVinco5" name="itemVinco5">
                            </div>

                            <div class="col-12 col-lg-4">
                                <label for="itemGramatura" class="form-label">Gramatura Bobina</label>
                                <select id="itemGramatura" name="itemGramatura" class="form-select form-select-sm select2">
                                    <option value="">Selecione...</option>
                                </select>
                            </div>

                            <div class="col-12 col-lg-4">
                                <label for="itemDiametroBobina" class="form-label">Diâmetro Bobina</label>
                                <input value="0" type="number" class="form-control form-control-sm" id="itemDiametroBobina" name="itemDiametroBobina">
                            </div>

                            <div class="col-12 col-lg-4">
                                <label for="itemLarguraBobina" class="form-label">Largura Bobina</label>
                                <input value="0" type="number" class="form-control form-control-sm" id="itemLarguraBobina" name="itemLarguraBobina">
                            </div>

                            <div class="col-12 col-lg-4">
                                <label for="itemTubete" class="form-label">Tubete Bobina</label>
                                <select class="form-select form-select-sm select2" name="itemTubete" id="itemTubete">
                                    <option value="">Selecione o tubete</option>
                                </select>
                            </div>

                            <div class="col-12 col-lg-4">
                                <label for="itemDataEntrega" class="form-label">Data de Entrega <span class="required-field">*</span></label>
                                <input type="date" class="form-control form-control-sm" id="itemDataEntrega" name="itemDataEntrega">
                            </div>

                            <div class="col-12 col-lg-4">
                                <label for="itemDataPrevistaEntrega" class="form-label">Data Prevista de Entrega </label>
                                <input type="date" class="form-control form-control-sm" id="itemDataPrevistaEntrega" name="itemDataPrevistaEntrega">
                            </div>
                            
                            <div class="col-12 col-lg-4 campos-bento">
                                <label for="itemModalFrete" class="form-label">Modal de frete</label>
                                <select id="itemModalFrete" name="itemModalFrete" class="form-select form-select-sm select2">
                                    <option value="">Selecione...</option>
                                    <option value="AEREO">AÉREO</option>
                                    <option value="MARITIMO">MARÍTIMO</option>
                                    <option value="RODOVIARIO">RODOVIÁRIO</option>
                                </select>
                            </div>

                            <div class="col-12 col-lg-4 campos-bento">
                                <label for="itemComercioExterior" class="form-label">Incoterms de Comércio Exterior</label>
                                <select id="itemComercioExterior" name="itemComercioExterior" class="form-select form-select-sm select2">
                                    <option value="" selected>Selecione...</option>
                                    <option value="EXW">EXW</option>
                                    <option value="FCA">FCA</option>
                                    <option value="FAS">FAS</option>
                                    <option value="FOB">FOB</option>
                                    <option value="CFR">CFR</option>
                                    <option value="CIF">CIF</option>
                                    <option value="DPU">DPU</option>
                                    <option value="CPT">CPT</option>
                                    <option value="DAP">DAP</option>
                                    <option value="CIP">CIP</option>
                                    <option value="DDP">DDP</option>
                                </select>
                            </div>

                            <div class="col-12 col-lg-4">
                                <label for="itemPedidoCompra" class="form-label">Pedido de Compra </label>
                                <input type="text" maxlength="15" class="form-control form-control-sm" id="itemPedidoCompra" name="itemPedidoCompra">
                            </div>

                            <div class="col-12 col-lg-4">
                                <label for="itemPedido" class="form-label">Item Pedido </label>
                                <input type="text" maxlength="15" class="form-control form-control-sm" id="itemPedido" name="itemPedido">
                            </div>

                            <div class="col-12">
                                <label for="itemObservacao" class="form-label">Observação</label>
                                <textarea class="form-control form-control-sm" id="itemObservacao" name="itemObservacao" rows="4"></textarea>
                            </div>
                        </div>
                    </form>
                </div>

                <div class="modal-footer">
                    <button type="button" class="btn btn-sm btn-outline-secondary" data-bs-dismiss="modal">
                        Cancelar
                    </button>
                    <button type="button" class="btn btn-sm btn-primary" id="btnSalvarItem">
                        <i class="bi bi-plus-lg me-1"></i>Adicionar
                    </button>
                </div>
            </div>
        </div>
    </div>

    <!-- Modal de Limite de Crédito do Cliente -->
    <div class="modal fade" id="modalLimiteCliente" tabindex="-1" aria-labelledby="modalLimiteClienteLabel" aria-hidden="true">
        <div class="modal-dialog modal-lg modal-dialog-centered">
            <div class="modal-content">
                <div class="modal-header">
                    <h5 class="modal-title" id="modalLimiteClienteLabel">
                        <i class="bi bi-cash-coin me-2"></i>Limite de Crédito do Cliente
                    </h5>
                    <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Fechar"></button>
                </div>
                <div class="modal-body">
                    <div id="limiteClienteConteudo"></div>
                </div>
            </div>
        </div>
    </div>

    <!-- Modal de Clientes -->
    <div class="modal fade" id="modalCliente"  aria-labelledby="modalClienteLabel" aria-hidden="true">
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

                    <div id="clienteResultados" style="max-height: 400px; overflow-y: auto;">
                    </div>
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


    <!-- Modal de ClienteRemessas Remessa -->
    <div class="modal fade" id="modalClienteRemessa"  aria-labelledby="modalClienteRemessaLabel" aria-hidden="true">
        <div class="modal-dialog modal-xl modal-dialog-centered modal-dialog-scrollable">
            <div class="modal-content">
                <div class="modal-header">
                    <h5 class="modal-title" id="modalClienteRemessaLabel">
                        <i class="bi bi-search me-2"></i>Buscar Cliente Remessa
                    </h5>
                    <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Fechar"></button>
                </div>
                <div class="modal-body">
                    <div class="mb-3">
                        <label for="clienteRemessaBusca" class="form-label">Pesquisar Cliente</label>
                        <div class="input-group">
                            <span class="input-group-text"><i class="bi bi-search"></i></span>
                            <input type="text" 
                                class="form-control form-control-sm" 
                                id="clienteRemessaBusca" 
                                placeholder="Digite o código, nome ou CNPJ para buscar..."
                                autocomplete="off">
                        </div>
                    </div>

                    <div id="clienteRemessaResultados" style="max-height: 400px; overflow-y: auto;">
                    </div>
                </div>
                <div class="modal-footer">
                    <button type="button" class="btn btn-sm btn-outline-secondary" data-bs-dismiss="modal">
                        Cancelar
                    </button>
                    <button type="button" class="btn btn-sm btn-primary" id="btnConfirmarClienteRemessa" onclick="confirmarClienteRemessa()" disabled>
                        <i class="bi bi-check-lg me-1"></i>Selecionar
                    </button>
                </div>
            </div>
        </div>
    </div>


    <!-- Modal de Produtos -->
    <div class="modal fade" id="modalProduto" aria-labelledby="modalProdutoLabel" aria-hidden="true">
        <div class="modal-dialog modal-xl modal-dialog-centered modal-dialog-scrollable">
            <div class="modal-content">
                <div class="modal-header">
                    <h5 class="modal-title" id="modalProdutoLabel">
                        <i class="bi bi-search me-2"></i>Buscar Produto
                    </h5>
                    <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Fechar"></button>
                </div>
                <div class="modal-body">
                    <div class="mb-3">
                        <label for="produtoBusca" class="form-label">Pesquisar Produto</label>
                        <div class="input-group">
                            <span class="input-group-text"><i class="bi bi-search"></i></span>
                            <input type="text" class="form-control form-control-sm" id="produtoBusca" placeholder="Digite o código, descrição ou referência para buscar..." autocomplete="off">
                        </div>
                    </div>
                    <div id="produtoResultados" style="max-height: 400px; overflow-y: auto;"></div>
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

    <!-- Modal de TES -->
    <div class="modal fade" id="modalTes" aria-labelledby="modalTesLabel" aria-hidden="true">
        <div class="modal-dialog modal-xl modal-dialog-centered modal-dialog-scrollable">
            <div class="modal-content">
                <div class="modal-header">
                    <h5 class="modal-title" id="modalTesLabel">
                        <i class="bi bi-search me-2"></i>Buscar TES
                    </h5>
                    <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Fechar"></button>
                </div>
                <div class="modal-body">
                    <div id="tesResultados" style="max-height: 400px; overflow-y: auto;">
                    </div>
                </div>
                <div class="modal-footer">
                    <button type="button" class="btn btn-sm btn-outline-secondary" data-bs-dismiss="modal">
                        Cancelar
                    </button>
                    <button type="button" class="btn btn-sm btn-primary" id="btnConfirmarTes" onclick="confirmarTes()" disabled>
                        <i class="bi bi-check-lg me-1"></i>Selecionar
                    </button>
                </div>
            </div>
        </div>
    </div>

</body>
</html>
