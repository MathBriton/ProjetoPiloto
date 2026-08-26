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
        .form-switch { display: flex; align-items: end; }
        .form-switch .form-check-input { width: 3em; height: 1.5em; cursor: pointer; }
        .form-switch .form-check-label { cursor: pointer; font-weight: 500; color: #495057; padding-left: 8px; }

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

        .anexo-item {
            border: 1px solid #e5e7eb;
            border-radius: 6px;
            margin-bottom: 10px;
            padding: 12px;
            transition: background-color .2s;
        }

        .anexo-item:hover {
            background-color: #f9fafb;
        }

        .anexo-icon {
            color: #3b82f6;
            font-size: 1.5rem;
            margin-right: 12px;
        }

        .anexo-info {
            flex: 1;
        }

        .anexo-nome {
            color: #1f2937;
            font-weight: 600;
            margin-bottom: 4px;
        }

        .anexo-data {
            color: #6b7280;
            font-size: .85rem;
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

    <script src="/corporativo/components/toast.js?v=<%=Now()%>"></script>
    <script src="/corporativo/components/confirmation.js?v=<%=Now()%>"></script>
    <script src="/corporativo/components/upload/fileUpload.js?v=<%=Now()%>"></script>

    <script>
        let id = 0;
        let abaAtual = 1;
        let totalAbas = 4;
        let abasValidadas = new Set();
        let validacoes = {};
        let acao = 'CRIAR';

        // Modal de Notas
        let listaNotas = [];
        let notasPesquisadas = [];
        let notaSelecionada = null;
        let indexNotaSelecionada = null;

        // Anexos
        let listaAnexos = [];

        function exibirLoading(exibir) {
            if (exibir) {
                $('#loadingOverlay').removeClass('d-none').addClass('d-flex');
            } else {
                $('#loadingOverlay').removeClass('d-flex').addClass('d-none');
            }
        }

        function normalizarAcao(valor) {
            const acaoNormalizada = (valor || '').toString().trim().toUpperCase();
            if (acaoNormalizada === 'EDICAO') return 'EDITAR';
            return ['CRIAR', 'EDITAR', 'VISUALIZAR'].includes(acaoNormalizada) ? acaoNormalizada : 'CRIAR';
        }
        function formularioSomenteVisualizacao() { 
            return acao === 'VISUALIZAR'; 
        }

        function formularioEmEdicao() { 
            return acao === 'EDITAR'; 
        }

        function notasObrigatorias() {
            return $('#tipoReclamacao').val() === 'FNC' || (formularioEmEdicao() && $('#flagConverterFNC').is(':checked'));
        }

        function configurarTipoReclamacao() {
            const emEdicao = formularioEmEdicao();
            $('#tipoReclamacao').prop('disabled', emEdicao || formularioSomenteVisualizacao()).trigger('change.select2');
            $('#linhaConverterFNC').toggleClass('d-none', !emEdicao);

            if (!emEdicao) $('#flagConverterFNC').prop('checked', false);

            $('#avisoConverterFNC').toggle(emEdicao && $('#flagConverterFNC').is(':checked'));
            $('#notaObrigatoria').toggle(notasObrigatorias());
        }

        function aplicarModoVisualizacao() {
            $('#formularioFNC input, #formularioFNC select, #formularioFNC textarea').prop('disabled', true);

            $('.select2').each(function() {
                $(this).prop('disabled', true).trigger('change.select2');
            });

            $('#divAcoes').removeClass('d-flex').addClass('d-none');
            $('#btnAdicionarNota').removeClass('d-flex').addClass('d-none');

            $('.tab-button').prop('disabled', false);
        }

        function inicializarSelect2() {
            $('.select2').select2({
                theme: 'bootstrap-5',
                width: '100%',
                dropdownCssClass: 'z-index-9999'
            });
        }

        function inicializarMascaras() {
            $('#quantidadeReclamada').mask('000.000.000.000.000', { reverse: true });
            $('#telefone').mask('(00) 00000-0000');
            $('#valorReclamado').maskMoney({
                prefix: '',
                thousands: '.',
                decimal: ',',
                allowZero: true,
                allowNegative: false
            });
        }

        function pesquisarCliente() {
            const codPlanta = $('#codPlanta').val();
            const baseUrl = `${location.protocol}//${location.host}`;
            const url = `${baseUrl}/corporativo/buscaCLIENTE.asp?id1=codCliente&id2=codLoja&id3=nomeCliente&id6=${codPlanta}&id9=cidadeEstado&id10=vendedor`;
            window.open(url, '', 'width=650,height=400');
        }

        function carregarDefeitos() {
            const $select = $('#defeitoId');
            $select.empty().append('<option value="">Todos os defeitos</option>');

            const params = {
                acao: 'listarDefeitos',
                tipoDefeito: $('#tipoDefeito').val()
            };

            $.ajax({
                url: 'FNCRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: params,
                dataType: 'json',
                success: function(response) {
                    if (!response.sucesso) {
                        console.error('Erro ao carregar defeitos: ' + response.mensagem);
                        return;
                    }

                    response.dados.forEach(function(item) {
                        $select.append($('<option>', { value: item.id, text: item.descricao }));
                    });
                },
                error: function(xhr, status, error) {
                    console.error('Erro ao carregar defeitos:', error);
                }
            });
        }

        function carregarDisposicoes() {
            const $select = $('#disposicaoId');
            $select.empty().append('<option value="">Todas as disposições</option>');

            const params = {
                acao: 'listarDisposicoes'
            };

            $.ajax({
                url: 'FNCRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: params,
                dataType: 'json',
                success: function(response) {
                    if (!response.sucesso) {
                        console.error('Erro ao carregar disposições: ' + response.mensagem);
                        return;
                    }

                    response.dados.forEach(function(item) {
                        $select.append($('<option>', { value: item.id, text: item.descricao }));
                    });
                },
                error: function(xhr, status, error) {
                    console.error('Erro ao carregar disposições:', error);
                }
            });
        }

        function pesquisarNotas() {
            const textoPesquisa = $('#campoPesquisa').val().trim();

            if (!textoPesquisa) {
                $('#resultadosPesquisa').html(`
                    <div class="text-center text-muted">
                        <i class="bi bi-search" style="font-size: 1.5rem;"></i>
                        <p class="mt-2 mb-0">Digite um termo para buscar notas existentes.</p>
                    </div>
                `);
                return;
            }

            notaSelecionada = null;
            indexNotaSelecionada = null;
            $('#btnsalvarNota').prop('disabled', true);

            $('#resultadosPesquisa').html(`
                <div class="text-center">
                    <div class="spinner-border spinner-border-sm text-primary" role="status">
                        <span class="visually-hidden">Buscando...</span>
                    </div>
                    <p class="mt-2 mb-0 text-muted">Buscando Notas...</p>
                </div>
            `);

            $.ajax({
                url: 'FNCRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: {
                    acao: 'listarNotasCliente',
                    codPlanta: $('#codPlanta').val(),
                    codCliente: $('#codCliente').val(),
                    codLoja: $('#codLoja').val(),
                    pesquisa: textoPesquisa
                },
                dataType: 'json',
                success: function(response) {
                    if (!response.sucesso) {
                        $('#resultadosPesquisa').html(`
                            <div class="alert text-center alert-danger" role="alert">
                                <i class="bi bi-exclamation-triangle" style="font-size: 1.5rem;"></i>
                                <p class="mt-2 mb-0"> ${response.mensagem || 'Erro ao buscar notas.'} </p>
                            </div>
                        `);
                        return;
                    }

                    notasPesquisadas = response.dados || [];
                    if (notasPesquisadas.length === 0) {
                        $('#resultadosPesquisa').html(`
                            <div class="alert text-center alert-info" role="alert">
                                <i class="bi bi-info-circle" style="font-size: 1.5rem;"></i>
                                <p class="mt-2 mb-0"> Nenhuma nota encontrada com os termos pesquisados. </p>
                            </div>
                        `);
                        return;
                    }

                    renderizarResultadosPesquisa();
                },
                error: function(xhr, status, error) {
                    console.error('Erro ao buscar notas:', error);
                    $('#resultadosPesquisa').html(`
                        <div class="alert text-center alert-info" role="alert">
                            <i class="bi bi-info-circle" style="font-size: 1.5rem;"></i>
                            <p class="mt-2 mb-0"> Nenhuma nota encontrada com os termos pesquisados.</p>
                        </div>
                    `);
                }
            });
        }

        function renderizarResultadosPesquisa() {
            let html = `
                <div style="overflow: auto; max-height: 45vh;" class="table-responsive">
                    <table class="table table-sm table-hover mb-0">
                        <thead class="thead-light">
                            <tr>
                                <th style="width: 50px;"></th>
                                <th>Nº Pedido</th>
                                <th>Item</th>
                                <th>Produto</th>
                                <th>Referência</th>
                                <th>Grupo</th>
                                <th>Quantidade</th>
                                <th>Preço Unit. NET</th>
                                <th>NF</th>
                                <th>Série</th>
                                <th>Emissão</th>
                            </tr>
                        </thead>
                        <tbody>
            `;

            notasPesquisadas.forEach((nota, index) => {
                html += `
                    <tr data-index="${index}">
                        <td class="text-center">
                            <button type="button" onclick="selecionarNota(${index})" class="btn-selecionar" title="Selecionar este item">
                                <i class="bi bi-circle"></i>
                            </button>
                        </td>
                        <td>${nota.numeroPedido}</td>
                        <td>${nota.itemPedido}</td>
                        <td>${nota.codProduto}</td>
                        <td>${nota.referencia}</td>
                        <td>${nota.grupoProduto}</td>
                        <td>${nota.quantidade}</td>
                        <td>${nota.precoUnitario}</td>
                        <td>${nota.notaFiscal}</td>
                        <td>${nota.serieNotaFiscal}</td>
                        <td>${nota.dataEmissaoNota}</td>
                    </tr>
                `;
            });

            html += `
                        </tbody>
                    </table>
                </div>
            `;

            $('#resultadosPesquisa').html(html);
        }

        function selecionarNota(index) {
            notaSelecionada = notasPesquisadas[index];
            indexNotaSelecionada = index;

            $('#resultadosPesquisa .btn-selecionar')
                .removeClass('selecionado')
                .html('<i class="bi bi-circle"></i>')
                .attr('title', 'Selecionar este item');

            $('#resultadosPesquisa tr[data-index="' + index + '"] .btn-selecionar')
                .addClass('selecionado')
                .html('<i class="bi bi-check-circle-fill"></i>')
                .attr('title', 'Item Selecionado');

            $('#btnsalvarNota').prop('disabled', false);
        }

        function abrirModalNota() {
            const codCliente = $('#codCliente').val();
            if (!codCliente) {
                Toast.warning('Informe o código do cliente antes de adicionar uma nota');
                return;
            }

            notaSelecionada = null;
            $('#campoPesquisa').val('');
            $('#resultadosPesquisa').html(`
                <div class="text-center text-muted">
                    <i class="bi bi-search" style="font-size: 1.5rem;"></i>
                    <p class="mt-2 mb-0">Digite um termo para buscar notas existentes.</p>
                </div>
            `);

            $('#modalNotaTitulo').html('<i class="bi bi-receipt me-2"></i>Adicionar Nota');
            $('#btnsalvarNota').prop('disabled', true);

            const modal = new bootstrap.Modal(document.getElementById('modalNota'));
            modal.show();

            setTimeout(function() {
                $('#campoPesquisa').focus();
            }, 500);
        }

        function removerNota(index) {
            ConfirmationModal.danger({
                message: 'Deseja realmente remover essa Nota?',
                type: 'danger',
                confirmText: 'Remover',
                cancelText: 'Cancelar',
                onConfirm: () => {
                    listaNotas.splice(index, 1);
                    renderizarNotas();
                    Toast.success('Nota removida com sucesso!');
                },
                onCancel: () => {}
            });
        }

        function salvarNota() {
            if (!notaSelecionada) {
                Toast.error('Selecione uma nota para adicionar');
                return;
            }

            if (listaNotas.some(nota => (nota.notaFiscal + nota.numeroPedido) === (notaSelecionada.notaFiscal + notaSelecionada.numeroPedido))) {
                Toast.error('Esta nota ja foi adicionada');
                return;
            }

            listaNotas.push({
                ...notaSelecionada,
                quantidadeReclamada: '',
                valorReclamado: '0,00'
            });
            Toast.success('Nota adicionada com sucesso!');
            renderizarNotas();

            const modal = bootstrap.Modal.getInstance(document.getElementById('modalNota'));
            modal.hide();
        }

        function obterQuantidadeReclamadaNumerica(valor) {
            return parseInt(String(valor || '').replace(/\D/g, ''), 10) || 0;
        }

        function formatarQuantidadeReclamada(valor) {
            const numero = String(valor === 0 ? 0 : (valor || '')).replace(/\D/g, '');
            return numero.replace(/\B(?=(\d{3})+(?!\d))/g, '.');
        }

        function obterValorNumerico(valor) {
            valor = String(valor || '0').replace(/[^\d,.-]/g, '');
            if (valor.indexOf(',') >= 0) {
                valor = valor.replace(/\./g, '').replace(',', '.');
            }
            return parseFloat(valor) || 0;
        }

        function formatarValorReclamado(valor) {
            return (parseFloat(valor) || 0).toLocaleString('pt-BR', {
                minimumFractionDigits: 2,
                maximumFractionDigits: 2
            });
        }

        function calcularValorReclamadoNota(nota) {
            return obterValorNumerico(nota.precoUnitario) * obterQuantidadeReclamadaNumerica(nota.quantidadeReclamada);
        }

        function calcularTotalQuantidadeReclamada() {
            return listaNotas.reduce((total, nota) => {
                return total + obterQuantidadeReclamadaNumerica(nota.quantidadeReclamada);
            }, 0);
        }

        function calcularTotalValorReclamado() {
            return listaNotas.reduce((total, nota) => {
                return total + calcularValorReclamadoNota(nota);
            }, 0);
        }

        function atualizarTotalQuantidadeReclamada() {
            const totalQuantidade = calcularTotalQuantidadeReclamada();
            const quantidadeFormatada = formatarQuantidadeReclamada(totalQuantidade) || '0';
            $('#quantidadeReclamada').val(quantidadeFormatada);
        }

        function atualizarValorReclamadoTotal(valor) {
            const valorFormatado = formatarValorReclamado(obterValorNumerico(valor));
            $('#valorReclamado').val(valorFormatado);
        }

        function atualizarTotaisReclamados() {
            atualizarTotalQuantidadeReclamada();
            atualizarValorReclamadoTotal(calcularTotalValorReclamado());
        }

        function atualizarQuantidadeReclamadaNota(index, valor) {
            const quantidadeFormatada = formatarQuantidadeReclamada(valor);
            listaNotas[index].quantidadeReclamada = quantidadeFormatada;
            listaNotas[index].valorReclamado = formatarValorReclamado(calcularValorReclamadoNota(listaNotas[index]));
            $(`.quantidade-reclamada-nota[data-index="${index}"]`).val(quantidadeFormatada);
            $(`.valor-reclamado-nota[data-index="${index}"]`).text(listaNotas[index].valorReclamado);
            atualizarTotaisReclamados();
        }

        function notasComQuantidadeReclamada() {
            return listaNotas.every(nota => String(nota.quantidadeReclamada || '').trim() !== '');
        }

        function renderizarNotas() {
            const $container = $('#notasListaContainer');
            $('#contadorNotas, #menuContadorNotas').text(listaNotas.length);

            if (listaNotas.length === 0) {
                $container.html(``);
                atualizarTotaisReclamados();
                return;
            }

            let html = `
                <div class="table-responsive">
                    <table class="table table-hover table-sm mb-0">
                        <thead class="table-light">
                            <tr>
                                <th>Nº do Pedido</th>
                                <th>Item</th>
                                <th>Referência</th>
                                <th>Quantidade Itens</th>
                                <th>Quantidade Reclamada</th>
                                <th class="text-end">Preço Unit.</th>
                                <th class="text-end">Valor Reclamado</th>
                                <th>Emissão</th>
                                <th>Nº Nota</th>
                                <th>Série</th>
                                <th class="text-center" style="width: 80px;">Ações</th>
                            </tr>
                        </thead>
                        <tbody>
            `;

            listaNotas.forEach((nota, index) => {
                nota.valorReclamado = formatarValorReclamado(calcularValorReclamadoNota(nota));
                
                html += `
                    <tr>
                        <td>${nota.numeroPedido}</td>
                        <td>${nota.itemPedido}</td>
                        <td>${nota.referencia}</td>
                        <td>${nota.quantidade}</td>
                        <td>
                            ${formularioSomenteVisualizacao() ? (nota.quantidadeReclamada || '') : `
                            <input type="text" class="form-control form-control-sm quantidade-reclamada-nota" data-index="${index}" value="${nota.quantidadeReclamada || ''}" oninput="atualizarQuantidadeReclamadaNota(${index}, this.value)">`}
                        </td>
                        <td class="text-end">${nota.precoUnitario}</td>
                        <td class="text-end valor-reclamado-nota" data-index="${index}">${nota.valorReclamado}</td>
                        <td>${nota.dataEmissaoNota}</td>
                        <td>${nota.notaFiscal}</td>
                        <td>${nota.serieNotaFiscal}</td>
                        <td class="text-center">
                            ${formularioSomenteVisualizacao() ? '' : `
                            <button type="button" class="btn btn-sm btn-danger" onclick="removerNota(${index})" title="Remover">
                                <i class="bi bi-trash"></i>
                            </button>`}
                        </td>
                    </tr>
                `;
            });

            html += `
                        </tbody>
                    </table>
                </div>
            `;

            $container.html(html);
            atualizarTotaisReclamados();
        }

        function renderizarAnexos() {
            const $container = $('#anexosListaContainer');
            $('#menuContadorAnexos').text(listaAnexos.length);

            if (listaAnexos.length === 0) {
                $container.html(``);
                return;
            }

            let html = '<div class="anexos-list mt-4">';

            listaAnexos.forEach((anexo, index) => {
                html += `
                    <div class="anexo-item d-flex align-items-center">
                        <i class="bi bi-file-earmark anexo-icon"></i>
                        <div class="anexo-info">
                            <div class="anexo-nome">${anexo.nome}</div>
                            <div class="anexo-data">Adicionado por <strong>${anexo.usuario}</strong> em: ${anexo.dataCriacao}</div>
                        </div>
                        <a href="/Clientes_Docs/${anexo.caminho}" target="_blank" class="btn btn-sm btn-primary">
                            <i class="bi bi-box-arrow-up-right me-1"></i>Abrir
                        </a>
                    </div>
                `;
            });

            html += '</div>';
            $container.html(html);
        }

        function inicializarValidacoes() {
            const naoVazio = (msg) => ({
                validator: (value) => value !== null && value !== undefined && String(value).trim() !== '',
                errorMessage: msg
            });

            const dataValida = (msg, { minDate, maxDate } = {}) => ({
                validator: (value) => {
                    const regex = /^\d{4}-(0[1-9]|1[0-2])-(0[1-9]|[12]\d|3[01])$/;
                    if (!regex.test(value)) return false;

                    const [year, month, day] = value.split('-').map(Number);
                    const date = new Date(year, month - 1, day);

                    const isValid =
                        date.getFullYear() === year &&
                        date.getMonth() === month - 1 &&
                        date.getDate() === day;

                    if (!isValid) return false;
                    if (minDate && date < minDate) return false;
                    if (maxDate && date > maxDate) return false;

                    return true;
                },
                errorMessage: msg
            });

            validacoes.aba1 = new JustValidate('#formularioFNC', { lockForm: false });
            validacoes.aba1
                .addField('#tipoReclamacao', [naoVazio('Tipo de reclamação é obrigatório')])
                .addField('#codPlanta', [naoVazio('Planta é obrigatória')])
                .addField('#codCliente', [naoVazio('Cliente é obrigatório')])
                .addField('#codLoja', [naoVazio('Loja é obrigatória')])
                .addField('#dataOcorrencia', [naoVazio('Data de ocorência é obrigatória'), dataValida('Data de ocorência inválida')])
                .addField('#tipoDefeito', [naoVazio('Tipo de defeito é obrigatório')])
                .addField('#defeitoId', [naoVazio('Classificação de defeito é obrigatória')])
                .addField('#disposicaoId', [naoVazio('Código Disposição')])
                .addField('#quantidadeReclamada', [naoVazio('Quantidade reclamada é obrigatória')])
                .addField('#valorReclamado', [naoVazio('Valor reclamado é obrigatório')])
                .addField('#descricaoDisposicao', [naoVazio('Disposição é obrigatória')])
                .addField('#descricaoDefeito', [naoVazio('Descrição do defeito é obrigatória')]);

            validacoes.aba2 = new JustValidate('#formularioFNC', { lockForm: false });
            validacoes.aba3 = new JustValidate('#formularioFNC', { lockForm: false });
            validacoes.aba4 = new JustValidate('#formularioFNC', { lockForm: false });
        }

        function exibirErrosValidacao(validator) {
            if (!validator || !validator.fields) {
                return;
            }

            const errorsParaExibir = {};

            Object.entries(validator.fields).forEach(([index, campo]) => {
                if (campo.isValid) {
                    return;
                }

                const primeiroErro = campo.rules.find(rule => !rule.isValid);

                if (primeiroErro && campo.elem && campo.elem.id) {
                    errorsParaExibir[`#${campo.elem.id}`] = primeiroErro.errorMessage || 'Campo inválido';
                }
            });

            validator.showErrors(errorsParaExibir);
        }

        function navegarParaGrade() {
            window.location.href = 'FNCGrade.asp';
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

        function liberarTodasAbas() {
            $('.tab-button').prop('disabled', false);
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

        function carregarDadosVisualizacao() {
            const params = {
                acao: 'obterPorId',
                id: id
            };

            $.ajax({
                url: 'FNCRepositorio.asp',
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
                    console.error('Erro ao carregar dados para visualização:', error);
                    Toast.error('Erro ao carregar dados para visualização');
                }
            });
        }

        function preencherFormulario(dados) {
            $('#tipoReclamacao').val(dados.tipoReclamacao).trigger('change.select2');
            $('#codPlanta').prop('disabled', true).val(dados.codPlanta).trigger('change.select2');
            $('#btnBuscarCliente').prop('disabled', true);
            $('#nomeCliente').prop('disabled', true).val(dados.nomeCliente);
            $('#codCliente').prop('disabled', true).val(dados.codCliente).trigger('change.select2');
            $('#codLoja').prop('disabled', true).val(dados.codLoja).trigger('change.select2');
            $('#tipoDefeito').val(dados.tipoDefeito).trigger('change.select2');
            $('#disposicaoId').val(dados.disposicaoId).trigger('change.select2');
            $('#quantidadeReclamada').val(dados.quantidadeReclamada);
            $('#valorReclamado').val(dados.valorReclamado);
            $('#descricaoDisposicao').val(dados.descricaoDisposicao);
            $('#descricaoDefeito').val(dados.descricaoDefeito);
            $('#contato').val(dados.contato);
            $('#email').val(dados.email);
            $('#telefone').val(dados.telefone);
            $('#dataOcorrencia').val(dados.dataOcorrencia);
            $('#cidadeEstado').val(dados.cidadeEstado);
            $('#vendedor').val(`${dados.codVendedor} - ${dados.nomeVendedor}`)

            listaNotas = dados.notas || [];
            listaAnexos = dados.anexos || [];
            renderizarNotas();
            atualizarValorReclamadoTotal(dados.valorReclamado);
            renderizarAnexos();
            configurarTipoReclamacao();

            var p1 = carregarDefeitos();

            $.when(p1).then(function() {
                setTimeout(() => {
                    $('#defeitoId').val(dados.defeitoId).trigger('change.select2');
                }, 500);
            });
        }

        function irProximo() {
            navegarParaAba(abaAtual + 1);
        }

        function irAnterior(){
            navegarParaAba(abaAtual - 1);
        }

        function anexarEventos() {
            $('#tipoDefeito').on('change', function() {
                carregarDefeitos();
            });

            $('#tipoReclamacao, #flagConverterFNC').on('change', configurarTipoReclamacao);

            $('#btnSalvar').on('click', function() {
                validarFormulario();
            });

            $('#formularioFNC').on('submit', function(e) {
                e.preventDefault();
                validarFormulario();
            });

            $('#formularioVinculoNota').on('submit', function(e) {
                e.preventDefault();
                pesquisarNotas();
            });

            $('#btnsalvarNota').on('click', function() {
                salvarNota();
            });

            $('#campoPesquisa').on('keypress', function(e) {
                if (e.which === 13) {
                    e.preventDefault();
                    $('#btnPesquisar').click();
                }
            });

            $('.tab-button').on('click', function() {
                const tabNumber = parseInt($(this).data('tab'));

                if (tabNumber === abaAtual) {
                    return;
                }

                navegarParaAba(tabNumber);
            });
        }

        async function validarFormulario() {
            if (formularioSomenteVisualizacao()) return;

            const abas = [1, 2, 3, 4];
            const notasValidas = !notasObrigatorias() || (listaNotas.length > 0 && notasComQuantidadeReclamada());
            const resultados = await Promise.all([
                validacoes.aba1.revalidate(),
                Promise.resolve(notasValidas),
                validacoes.aba3.revalidate(),
                validacoes.aba4.revalidate(),
            ]);

            resultados.forEach((valida, index) => {
                const numeroAba = abas[index];
                if (!valida) {
                    marcarAbaComoInvalida(numeroAba);
                    exibirErrosValidacao(validacoes[`aba${numeroAba}`]);
                } else {
                    marcarAbaComoCompleta(numeroAba);
                }
            });

            if (notasObrigatorias() && listaNotas.length === 0) {
                Toast.error('Adicione ao menos uma nota');
            } else if (notasObrigatorias() && !notasComQuantidadeReclamada()) {
                Toast.error('Informe a quantidade reclamada de todas as notas');
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
            atualizarTotalQuantidadeReclamada();
            atualizarValorReclamadoTotal($('#valorReclamado').val());

            const notas = listaNotas.map(nota => {
                return {
                    ...nota,
                    quantidadeReclamada: String(nota.quantidadeReclamada || '').replace(/\D/g, ''),
                    valorReclamado: obterValorNumerico(nota.valorReclamado).toFixed(2),
                    dataEmissaoNota: nota.dataEmissaoNota && nota.dataEmissaoNota.includes('/')
                        ? nota.dataEmissaoNota.split('/').reverse().join('-')
                        : nota.dataEmissaoNota
                };
            });

            const anexos = uploadAnexos ? uploadAnexos.getFiles() : [];

            const dadosFormulario = {
                acao: formularioEmEdicao() ? 'editarSolicitante' : 'criar',
                tipoReclamacao: $('#tipoReclamacao').val(),
                codPlanta: $('#codPlanta').val(),
                nomeCliente: $('#nomeCliente').val(),
                codCliente: $('#codCliente').val(),
                codLoja: $('#codLoja').val(),
                dataOcorrencia: $('#dataOcorrencia').val(),
                tipoDefeito: $('#tipoDefeito').val(),
                defeitoId: $('#defeitoId').val(),
                disposicaoId: $('#disposicaoId').val(),
                quantidadeReclamada: $('#quantidadeReclamada').val().replace(/\D/g, ''),
                valorReclamado: $('#valorReclamado').val().replace(/\./g, '').replace(/,/g, '.'),
                descricaoDisposicao: $('#descricaoDisposicao').val().replaceAll("'", '"'),
                descricaoDefeito: $('#descricaoDefeito').val().replaceAll("'", '"'),
                contato: $('#contato').val(),
                email: $('#email').val(),
                telefone: $('#telefone').val(),
                notas: JSON.stringify(notas),
                anexos: JSON.stringify(anexos)
            };

            if (id) {
                dadosFormulario.id = id;
                if (formularioEmEdicao() && $('#flagConverterFNC').is(':checked')) {
                    dadosFormulario.flagConverterFNC = '1';
                }
            }

            exibirLoading(true);

            $.ajax({
                url: 'FNCRepositorio.asp',
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

        function verificarModoFormulario() {
            const urlParams = new URLSearchParams(window.location.search);
            id = urlParams.get('id') || 0;
            acao = normalizarAcao(urlParams.get('acao'));
            if (id) {
                if (formularioSomenteVisualizacao()) $('#titulo').text(`Visualizar Cadastro de Reclamação # ${id}`);
                else { acao = 'EDITAR'; $('#titulo').text(`Editar Cadastro de Reclamação # ${id}`); }
            } else {
                acao = 'CRIAR';
                $('#titulo').text('Cadastro de Reclamação');
            }

            configurarTipoReclamacao();
        }

        $(document).ready(function() {
            anexarEventos();
            navegarParaAba(1);
            verificarModoFormulario();

            carregarDisposicoes();

            if (!formularioSomenteVisualizacao()) {
                inicializarValidacoes();
            }

            inicializarSelect2();
            inicializarMascaras();
            renderizarNotas();
            renderizarAnexos();

            if (id) {
                liberarTodasAbas();
                if (formularioSomenteVisualizacao()) aplicarModoVisualizacao();
                carregarDadosVisualizacao();
            }
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

        <div class="form-container">
            <div class="form-layout">
                <div class="tabs-container">
                    <div class="tabs">
                        <div class="tab-group-title">Geral</div>

                        <button type="button" class="tab-button active" data-tab="1">
                            <i class="tab-icon bi bi-clipboard-check"></i>
                            <span class="tab-title">Cadastro</span>
                        </button>

                        <button type="button" class="tab-button" data-tab="2">
                            <i class="tab-icon bi bi-receipt"></i>
                            <span class="tab-title">Notas</span>
                            <span class="tab-meta" id="menuContadorNotas">0</span>
                        </button>

                        <button type="button" class="tab-button" data-tab="3">
                            <i class="tab-icon bi bi-person-lines-fill"></i>
                            <span class="tab-title">Contato</span>
                        </button>

                        <div class="tab-group-title">Adicionais</div>

                        <button type="button" class="tab-button" data-tab="4">
                            <i class="tab-icon bi bi-paperclip"></i>
                            <span class="tab-title">Anexos</span>
                            <span class="tab-meta" id="menuContadorAnexos">0</span>
                        </button>
                    </div>
                </div>

                <div class="tab-content-wrapper">
                    <form id="formularioFNC" name="formularioFNC" novalidate>
                        <div class="tab-content active" data-tab-content="1">
                            <div class="section-header">
                                <div class="section-title">Cadastro da reclamação</div>
                                <p class="section-subtitle">Dados principais da reclamação.</p>

                                <div id="linhaConverterFNC" class="my-2 d-none">
                                    <div class="form-check form-switch">
                                        <input class="form-check-input" type="checkbox" id="flagConverterFNC" name="flagConverterFNC" value="1">
                                        <label class="form-check-label" for="flagConverterFNC">Converter em FNC</label>
                                    </div>
                                </div>
                            </div>

                            <div id="avisoConverterFNC" class="alert alert-primary alert-dismissible py-2 px-3 mb-3 align-items-center" role="alert" style="font-size: .85rem; display: none;">
                                <i class="bi bi-info-circle me-2"></i>Para converter em FNC, adicione ao menos uma nota antes de salvar.
                            </div>

                            <div class="row g-3">
                                <div class="col-12 col-md-6">
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

                                <div class="col-12 col-md-6">
                                    <label for="tipoReclamacao" class="form-label">Tipo de Reclamação <span class="required-field">*</span></label>
                                    <select class="form-select form-select-sm select2" name="tipoReclamacao" id="tipoReclamacao">
                                        <option value="">Selecione um tipo</option>
                                        <option value="NOTIFICACAO">NOTIFICACAO</option>
                                        <option value="FNC">FNC</option>
                                    </select>
                                </div>

                                <div class="col-12 col-md-2">
                                    <label class="form-label">Buscar cliente</label>
                                    <div class="input-group">
                                        <button id="btnBuscarCliente" type="button" class="flex-grow-1 btn btn-sm btn-secondary" onclick="pesquisarCliente()">
                                            Buscar <i class="bi bi-box-arrow-up-right"></i>
                                        </button>
                                    </div>
                                </div>

                                <div class="col-12 col-md-3">
                                    <label for="nomeCliente" class="form-label">Nome do Cliente <span class="required-field">*</span></label>
                                    <input type="text" class="form-control form-control-sm" id="nomeCliente" name="nomeCliente" disabled>
                                </div>

                                <div class="col-12 col-md-4">
                                    <label for="cidadeEstado" class="form-label">Cidade/Estado <span class="required-field">*</span></label>
                                    <input type="text" class="form-control form-control-sm" id="cidadeEstado" name="cidadeEstado" disabled>
                                </div>

                                <div class="col-12 col-md-2">
                                    <label for="codCliente" class="form-label">Cód. Cliente <span class="required-field">*</span></label>
                                    <input disabled type="text" class="form-control form-control-sm" id="codCliente" name="codCliente">
                                </div>

                                <div class="col-12 col-md-1">
                                    <label for="codLoja" class="form-label">Loja <span class="required-field">*</span></label>
                                    <input type="text" class="form-control form-control-sm" id="codLoja" name="codLoja" disabled>
                                </div>

                                <div class="col-12 col-md-4">
                                    <label for="vendedor" class="form-label">Vendedor <span class="required-field">*</span></label>
                                    <input type="text" class="form-control form-control-sm" id="vendedor" name="vendedor" disabled>
                                </div>

                                <div class="col-12 col-md-4">
                                    <label for="dataOcorrencia" class="form-label">Data de Ocorrência <span class="required-field">*</span></label>
                                    <input type="date" class="form-control form-control-sm" id="dataOcorrencia" name="dataOcorrencia">
                                </div>

                                <div class="col-12 col-md-4">
                                    <label for="tipoDefeito" class="form-label">Tipo de Defeito <span class="required-field">*</span></label>
                                    <select class="form-select form-select-sm select2" name="tipoDefeito" id="tipoDefeito">
                                        <option value="">Selecione a tipo de defeito</option>
                                        <option value="ADM">Administrativo</option>
                                        <option value="PRO">Industrial</option>
                                    </select>
                                </div>

                                <div class="col-12 col-md-4">
                                    <label for="defeitoId" class="form-label">Classificação Defeito <span class="required-field">*</span></label>
                                    <select class="form-select form-select-sm select2" name="defeitoId" id="defeitoId">
                                        <option value="">Selecione um defeito</option>
                                    </select>
                                </div>

                                <div class="col-12 col-md-4">
                                    <label for="disposicaoId" class="form-label">Código Disposição <span class="required-field">*</span></label>
                                    <select class="form-select form-select-sm select2" name="disposicaoId" id="disposicaoId">
                                        <option value="">Selecione a disposição</option>
                                    </select>
                                </div>

                                <div class="col-12">
                                    <label for="descricaoDisposicao" class="form-label">Disposição <span class="required-field">*</span></label>
                                    <textarea class="form-control form-control-sm" id="descricaoDisposicao" name="descricaoDisposicao" rows="4"></textarea>
                                </div>

                                <div class="col-12">
                                    <label for="descricaoDefeito" class="form-label">Descrição Defeito <span class="required-field">*</span></label>
                                    <textarea class="form-control form-control-sm" id="descricaoDefeito" name="descricaoDefeito" rows="4"></textarea>
                                </div>
                            </div>
                        </div>
                        
                        <div class="tab-content" data-tab-content="2">
                            <div class="section-header d-flex justify-content-between align-items-center">
                                <div>
                                    <div class="section-title">
                                        Notas vinculadas <span id="notaObrigatoria" class="required-field" style="display:none;">*</span> <span class="badge bg-secondary ms-1" id="contadorNotas">0</span>
                                    </div>
                                    <p class="section-subtitle">Notas fiscais relacionadas à reclamação.</p>
                                </div>

                                <button type="button" class="btn btn-sm btn-primary" id="btnAdicionarNota" onclick="abrirModalNota()">
                                    <i class="bi bi-plus-lg me-1"></i>Adicionar
                                </button>
                            </div>

                            <div class="row g-3 mb-3">
                                <div class="col-lg-4 col-md-6">
                                    <label for="quantidadeReclamada" class="form-label">Quantidade Reclamada Total</label>
                                    <input type="text" class="form-control form-control-sm" id="quantidadeReclamada" value="0" readonly>
                                </div>
                                <div class="col-lg-4 col-md-6">
                                    <label for="valorReclamado" class="form-label">Valor Reclamado Total</label>
                                    <input type="text" class="form-control form-control-sm" id="valorReclamado" value="0,00" >
                                </div>
                            </div>
                            <div id="notasListaContainer"></div>
                        </div>
                        
                        <div class="tab-content" data-tab-content="3">
                            <div class="section-header">
                                <div class="section-title">Contato</div>
                                <p class="section-subtitle">Dados do responsável pelo recebimento ou retorno ao cliente.</p>
                            </div>

                            <div class="row g-3">
                                <div class="col-lg-4 col-md-12">
                                    <label for="contato" class="form-label">Contato</label>
                                    <input type="text" class="form-control form-control-sm" id="contato" name="contato" placeholder="Nome do responsável pelo recebimento">
                                </div>

                                <div class="col-lg-6 col-md-12">
                                    <label for="email" class="form-label">E-mail</label>
                                    <input type="email" class="form-control form-control-sm" id="email" name="email" placeholder="email@exemplo.com.br">
                                </div>

                                <div class="col-lg-2 col-md-12">
                                    <label for="telefone" class="form-label">Telefone</label>
                                    <input type="text" class="form-control form-control-sm" id="telefone" name="telefone" placeholder="(00) 00000-0000" maxlength="15">
                                </div>
                            </div>
                        </div>

                        <div class="tab-content" data-tab-content="4">
                            <div class="section-header">
                                <div class="section-title">Anexos</div>
                                <p class="section-subtitle">Evidências, documentos e arquivos relacionados à FNC.</p>
                            </div>

                            <div class="row g-3">
                                <file-upload id="uploadAnexos" file-code="FNC" max-files="10" multiple></file-upload>
                                <div id="anexosListaContainer"></div>
                            </div>
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

    <div class="modal fade" id="modalNota" tabindex="-1" aria-labelledby="modalNotaLabel" aria-hidden="true">
        <div class="modal-dialog modal-xl">
            <div class="modal-content">
                <div class="modal-header">
                    <h5 class="modal-title" id="modalNotaTitulo">
                        <i class="bi bi-receipt me-2"></i>Adicionar Nota
                    </h5>
                    <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Fechar"></button>
                </div>
                <div class="modal-body">
                    <form id="formularioVinculoNota" name="formularioVinculoNota" novalidate>
                        <div class="row mb-3">
                            <div class="col-md-10">
                                <label for="campoPesquisa" class="form-label">Pesquisar Nota <span class="required-field">*</span></label>
                                <input required type="text" class="form-control form-control-sm" id="campoPesquisa" name="campoPesquisa" placeholder="Digite o número da nota, pedido ou produto...">
                            </div>
                            <div class="col-md-2 d-flex align-items-end">
                                <button type="submit" class="btn btn-sm btn-primary w-100" id="btnPesquisar">
                                    <i class="bi bi-search me-1"></i>Pesquisar
                                </button>
                            </div>
                        </div>
                    </form>

                    <div id="resultadosPesquisa"></div>
                </div>
                <div class="modal-footer">
                    <button type="button" class="btn btn-sm btn-secondary" data-bs-dismiss="modal">Cancelar</button>
                    <button type="button" class="btn btn-sm btn-primary" id="btnsalvarNota" disabled>
                        <i class="bi bi-plus-lg me-1"></i>Adicionar
                    </button>
                </div>
            </div>
        </div>
    </div>
</body>
</html>
