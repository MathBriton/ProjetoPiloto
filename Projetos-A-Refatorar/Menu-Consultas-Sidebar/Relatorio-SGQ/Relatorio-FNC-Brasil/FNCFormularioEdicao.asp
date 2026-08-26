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
    <link href="../Diagramas/Ishikawa/ishikawa.css?v=<%=Now()%>" rel="stylesheet">
    <link href="../Diagramas/5W2H/5W2H.css?v=<%=Now()%>" rel="stylesheet">

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

        #btnVoltarPagina:disabled,
        #btnProximaPagina:disabled {
            cursor: not-allowed;
            pointer-events: auto;
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

        /* Layout principal */
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
            display: flex;
            flex-direction: column;
            background: #fff;
            min-width: 0;
            justify-content: space-between;
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

        th,
        td {
            white-space: nowrap;
            vertical-align: top;
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

        .btn-selecionar:disabled {
            cursor: not-allowed;
            opacity: .5;
        }

        .btn-selecionar.selecionado {
            background-color: #198754;
            border-color: #198754;
            color: #fff;
            cursor: default;
        }

        .btn-selecionar.selecionado:hover {
            background-color: #198754;
            border-color: #198754;
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

        .nqa-inline {
            animation: fadeIn .2s ease;
            border-bottom: 1px dashed #d5d9e2;
            display: none;
            margin-bottom: 14px;
            padding-bottom: 12px;
        }

        .nqa-inline.visible {
            align-items: center;
            display: flex;
            flex-wrap: wrap;
            gap: 16px;
        }

        .nq-row-label {
            color: #6b7283;
            font-size: 12px;
            white-space: nowrap;
        }

        .nq-val {
            font-family: monospace;
            font-size: 15px;
            font-weight: 500;
            transition: color .3s;
            white-space: nowrap;
        }

        .nq-sep {
            background: #d5d9e2;
            flex-shrink: 0;
            height: 14px;
            width: 1px;
        }

        .nq-badge {
            border-radius: 3px;
            font-size: 10px;
            font-weight: 600;
            letter-spacing: .7px;
            padding: 2px 9px;
            text-transform: uppercase;
            transition: all .3s;
            white-space: nowrap;
        }

        .nq-detail {
            color: #6b7283;
            font-size: 11px;
        }

        .nqa-inline.s-ok .nq-val { color: #0f7b4f; }
        .nqa-inline.s-ok .nq-badge { background: #e6f5ee; color: #0f7b4f; border: 1px solid #b3dfc8; }
        .nqa-inline.s-warn .nq-val { color: #8a6200; }
        .nqa-inline.s-warn .nq-badge { background: #fef8e7; color: #8a6200; border: 1px solid #f0d880; }
        .nqa-inline.s-fail .nq-val { color: #b91c1c; }
        .nqa-inline.s-fail .nq-badge { background: #fef2f2; color: #b91c1c; border: 1px solid #fca5a5; }


        .truncate-text {
            display: -webkit-box;
            -webkit-line-clamp: 2;
            -webkit-box-orient: vertical;
            max-width: 360px;
            overflow: hidden;
            text-overflow: ellipsis;
            white-space: normal;
            word-break: break-word;
        }

        .texto-detalhe-cell {
            min-width: 260px;
            vertical-align: top;
        }

        .texto-detalhe-link {
            cursor: pointer;
            transition: color .15s ease;
        }

        .texto-detalhe-link:hover,
        .texto-detalhe-link:focus {
            color: #0d6efd;
            outline: none;
            text-decoration: underline;
        }

        .detalhe-conteudo {
            background: #f8f9fa;
            border: 1px solid #e9ecef;
            border-radius: 8px;
            max-height: 60vh;
            overflow: auto;
            padding: 14px;
            white-space: pre-wrap;
            word-break: break-word;
        }

        @media (max-width: 800px) {
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
    <script src="../Diagramas/Ishikawa/ishikawa.js?v=<%=Now()%>"></script>
    <script src="../Diagramas/5W2H/5W2H.js?v=<%=Now()%>"></script>
    <script src="/corporativo/funcoes/funcoesJAVASCRIPT.js?v=<%=Now()%>"></script>

    <script>
        let id = 0;
        let abaAtual = 1;
        let totalAbas = 10;
        let abasValidadas = new Set();
        let validacoes = {};
        let nivelUsuario = '';
        let formularioSomenteAnexo = false;

        // Modal de Notas
        let listaNotas = [];
        let notasPesquisadas = [];
        let notaSelecionada = null;

        // Auxiliares
        let listaDefeitos = [];
        let listaAnexos = [];
        let listaObservacoes = [];
        let listaRespostas = [];
        let listaAprovacoes = [];
        let limitesNQA = {};


        const eSituacao = Object.freeze({
            ABERTA: '1',
            RESPONDIDA: '2',
            ENCERRADA: '3'
        });

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

        function repoUrl() {
            return 'FNCRepositorio.asp?v=<%=Now()%>';
        }

        function popularSelect($select, response, mapFn) {
            if (!response.sucesso) {
                console.error(response.mensagem);
                return;
            }

            response.dados.forEach(function(item) {
                $select.append($('<option>', mapFn(item)));
            });
        }


        function escapeHtml(valor) {
            return String(valor || '')
                .replace(/&/g, '&amp;')
                .replace(/</g, '&lt;')
                .replace(/>/g, '&gt;')
                .replace(/"/g, '&quot;')
                .replace(/'/g, '&#39;');
        }

        function normalizarTextoDetalhe(valor) {
            const texto = String(valor || '').trim();
            return texto || '-';
        }

        function criarCelulaDetalhe(valor, titulo) {
            const texto = normalizarTextoDetalhe(valor);
            const textoSeguro = escapeHtml(texto);
            const tituloSeguro = escapeHtml(titulo || 'Detalhes');

            if (texto === '-') {
                return '<td class="texto-detalhe-cell text-muted">-</td>';
            }

            return `
                <td class="texto-detalhe-cell">
                    <div class="truncate-text texto-detalhe-link"
                         role="button"
                         tabindex="0"
                         onclick="abrirModalDetalhe(this)"
                         onkeydown="abrirModalDetalheTeclado(event, this)"
                         data-titulo="${tituloSeguro}"
                         data-detalhe="${textoSeguro}">
                        ${textoSeguro}
                    </div>
                </td>
            `;
        }

        function abrirModalDetalhe(botao) {
            const titulo = $(botao).attr('data-titulo') || 'Detalhes';
            const detalhe = $(botao).attr('data-detalhe') || '-';

            $('#modalDetalheTituloTexto').text(titulo);
            $('#modalDetalheConteudo').text(detalhe);

            const modal = new bootstrap.Modal(document.getElementById('modalDetalheTexto'));
            modal.show();
        }

        function abrirModalDetalheTeclado(event, elemento) {
            if (event.key === 'Enter' || event.key === ' ') {
                event.preventDefault();
                abrirModalDetalhe(elemento);
            }
        }

        function carregarTiposDefeito() {
            const $select = $('#tipoDefeito');
            $select.empty().append('<option value="">Selecione o tipo de defeito</option>');

            return $.ajax({
                url: 'FNCRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: { acao: 'listarTiposDefeito' },
                dataType: 'json',
                success: function(response) {
                    popularSelect($select, response, item => ({ value: item.id, text: item.descricao }));
                },
                error: function(xhr, status, error) {
                    console.error('Erro ao carregar tipos de defeito:', error);
                }
            });
        }

        function carregarSituacoes() {
            const $select = $('#situacao');
            $select.empty().append('<option value="">Selecione a situação</option>');

            return $.ajax({
                url: repoUrl(),
                type: 'GET',
                data: { acao: 'listarSituacoes' },
                dataType: 'json',
                success: function(response) {
                    popularSelect($select, response, item => ({ value: item.id, text: item.descricao }));
                },
                error: function(xhr, status, error) {
                    console.error('Erro ao carregar situações:', error);
                }
            });
        }

        function carregarStatus() {
            const $select = $('#status');
            $select.empty().append('<option value="">Selecione o status</option>');

            return $.ajax({
                url: repoUrl(),
                type: 'GET',
                data: { acao: 'listarStatus' },
                dataType: 'json',
                success: function(response) {
                    popularSelect($select, response, item => ({ value: item.id, text: item.descricao }));
                },
                error: function(xhr, status, error) {
                    console.error('Erro ao carregar status:', error);
                }
            });
        }

        function carregarSetores() {
            const $select = $('#setorId');
            $select.empty().append('<option value="">Selecione o setor</option>');

            return $.ajax({
                url: repoUrl(),
                type: 'GET',
                data: { acao: 'listarSetores' },
                dataType: 'json',
                success: function(response) {
                    popularSelect($select, response, item => ({ value: item.id, text: item.descricao }));
                },
                error: function(xhr, status, error) {
                    console.error('Erro ao carregar setores:', error);
                }
            });
        }

        function carregarDisposicoes() {
            const $select = $('#disposicaoId');
            $select.empty().append('<option value="">Selecione a disposição</option>');

            return $.ajax({
                url: repoUrl(),
                type: 'GET',
                data: { acao: 'listarDisposicoes' },
                dataType: 'json',
                success: function(response) {
                    popularSelect($select, response, item => ({ value: item.id, text: item.descricao }));
                },
                error: function(xhr, status, error) {
                    console.error('Erro ao carregar disposições:', error);
                }
            });
        }

        function carregarDefeitos(tipoDefeito) {
            const $select = $('#defeitoId');
            $select.empty().append('<option value="">Selecione um defeito</option>');

            if (!tipoDefeito) {
                return $.Deferred().resolve().promise();
            }

            return $.ajax({
                url: repoUrl(),
                type: 'GET',
                data: {
                    acao: 'listarDefeitos',
                    tipoDefeito: tipoDefeito
                },
                dataType: 'json',
                success: function(response) {
                    if (!response.sucesso) {
                        console.error('Erro ao carregar defeitos: ' + response.mensagem);
                        return;
                    }

                    response.dados.forEach(function(item) {
                        $select.append($('<option>', { value: item.id, text: item.descricao }));
                    });

                    listaDefeitos = response.dados;
                },
                error: function(xhr, status, error) {
                    console.error('Erro ao carregar defeitos:', error);
                }
            });
        }

        function carregarResponsaveis(codPlanta) {
            const $select = $('#responsavelId');
            $select.empty().append('<option value="">Selecione um responsável</option>');

            return $.ajax({
                url: repoUrl(),
                type: 'GET',
                data: {
                    acao: 'listarResponsaveis',
                    codPlanta: codPlanta
                },
                dataType: 'json',
                success: function(response) {
                    if (!response.sucesso) {
                        console.error('Erro ao carregar responsáveis: ' + response.mensagem);
                        return;
                    }

                    response.dados.forEach(function(item) {
                        $select.append($('<option>', { value: item.id, text: item.nome }));
                    });
                },
                error: function(xhr, status, error) {
                    console.error('Erro ao carregar responsáveis:', error);
                }
            });
        }

        function carregarLimitesNQA() {
            return $.ajax({
                url: repoUrl(),
                type: 'GET',
                data: { acao: 'obterLimitesNQA' },
                dataType: 'json',
                success: function(response) {
                    if (!response.sucesso) {
                        console.error('Erro ao carregar limites NQA: ' + response.mensagem);
                        return;
                    }

                    limitesNQA = response.dados;
                },
                error: function(xhr, status, error) {
                    console.error('Erro ao carregar limites NQA:', error);
                }
            });
        }

        function pesquisarNotas() {
            const textoPesquisa = $('#campoPesquisa').val().trim();

            if (!textoPesquisa) {
                $('#resultadosPesquisaNota').html(`
                    <div class="text-center text-muted">
                        <i class="bi bi-search" style="font-size: 1.5rem;"></i>
                        <p class="mt-2 mb-0">Digite um termo para buscar notas existentes.</p>
                    </div>
                `);
                return;
            }

            notaSelecionada = null;
            $('#btnSalvarNota').prop('disabled', true);

            $('#resultadosPesquisaNota').html(`
                <div class="text-center">
                    <div class="spinner-border spinner-border-sm text-primary" role="status">
                        <span class="visually-hidden">Buscando...</span>
                    </div>
                    <p class="mt-2 mb-0 text-muted">Buscando notas...</p>
                </div>
            `);

            $.ajax({
                url: repoUrl(),
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
                        $('#resultadosPesquisaNota').html(`
                            <div class="alert text-center alert-danger" role="alert">
                                ${response.mensagem || 'Erro ao buscar notas.'}
                            </div>
                        `);
                        return;
                    }

                    notasPesquisadas = response.dados || [];

                    if (notasPesquisadas.length === 0) {
                        $('#resultadosPesquisaNota').html(`
                            <div class="alert text-center alert-info" role="alert">
                                Nenhuma nota encontrada com os termos pesquisados.
                            </div>
                        `);
                        return;
                    }

                    renderizarResultadosPesquisaNotas();
                },
                error: function() {
                    $('#resultadosPesquisaNota').html(`
                        <div class="alert text-center alert-info" role="alert">
                            Nenhuma nota encontrada com os termos pesquisados.
                        </div>
                    `);
                }
            });
        }

        function renderizarResultadosPesquisaNotas() {
            let html = `
                <div style="overflow: auto; max-height: 45vh;" class="table-responsive">
                    <table class="table table-sm table-hover mb-0">
                        <thead class="table-light">
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

            $('#resultadosPesquisaNota').html(html);
        }

        function selecionarNota(index) {
            notaSelecionada = notasPesquisadas[index];

            $('#resultadosPesquisaNota .btn-selecionar')
                .removeClass('selecionado')
                .html('<i class="bi bi-circle"></i>')
                .attr('title', 'Selecionar este item');

            $('#resultadosPesquisaNota tr[data-index="' + index + '"] .btn-selecionar')
                .addClass('selecionado')
                .html('<i class="bi bi-check-circle-fill"></i>')
                .attr('title', 'Item selecionado');

            $('#btnSalvarNota').prop('disabled', false);
        }

        function abrirModalNota() {
            if (!$('#codCliente').val()) {
                Toast.warning('Informe o código do cliente antes de adicionar uma nota.');
                return;
            }

            notaSelecionada = null;
            notasPesquisadas = [];
            $('#campoPesquisa').val('');
            $('#btnSalvarNota').prop('disabled', true);
            $('#resultadosPesquisaNota').html(`
                <div class="text-center text-muted">
                    <i class="bi bi-search" style="font-size: 1.5rem;"></i>
                    <p class="mt-2 mb-0">Digite um termo para buscar notas existentes.</p>
                </div>
            `);

            const modal = new bootstrap.Modal(document.getElementById('modalNota'));
            modal.show();

            setTimeout(function() {
                $('#campoPesquisa').focus();
            }, 500);
        }

        function salvarNota() {
            if (!notaSelecionada) {
                Toast.error('Selecione uma nota para adicionar');
                return;
            }

            if (listaNotas.some(nota => (nota.notaFiscal + nota.numeroPedido) === (notaSelecionada.notaFiscal + notaSelecionada.numeroPedido))) {
                Toast.error('Esta nota já foi adicionada');
                return;
            }

            listaNotas.push({
                ...notaSelecionada,
                quantidadeReclamada: '',
                valorReclamado: '0,00'
            });
            Toast.success('Nota adicionada com sucesso!');

            renderizarNotas();
            calcularCoeficienteNQA();

            bootstrap.Modal.getInstance(document.getElementById('modalNota')).hide();
        }

        function removerNota(index) {
            ConfirmationModal.danger({
                message: 'Deseja realmente remover essa nota?',
                type: 'danger',
                confirmText: 'Remover',
                cancelText: 'Cancelar',
                onConfirm: () => {
                    listaNotas.splice(index, 1);
                    renderizarNotas();
                    calcularCoeficienteNQA();
                    Toast.success('Nota removida com sucesso!');
                },
                onCancel: () => {}
            });
        }

        function renderizarNotas() {  
            const $container = $('#notasListaContainer');  
            $('#contadorNotas, #menuContadorNotas').text(listaNotas.length);

            if (listaNotas.length === 0) {  
                $container.html(`  
                    <div class="text-center py-5">  
                        <p class="text-muted mt-3 mb-0">Nenhuma nota adicionada.</p>  
                    </div>  
                `);  
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
                                ${usuarioAnalista() ? '<th class="text-center" style="width: 80px;">Ações</th>' : ''}  
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
                            ${usuarioAnalista() ? `  
                            <input type="text" class="form-control form-control-sm quantidade-reclamada-nota" data-index="${index}" value="${nota.quantidadeReclamada || ''}" oninput="atualizarQuantidadeReclamadaNota(${index}, this.value)">` : (nota.quantidadeReclamada || '')}  
                        </td>  
                        <td class="text-end">${nota.precoUnitario}</td>  
                        <td class="text-end valor-reclamado-nota" data-index="${index}">${nota.valorReclamado}</td>  
                        <td>${nota.dataEmissaoNota}</td>  
                        <td>${nota.notaFiscal}</td>  
                        <td>${nota.serieNotaFiscal}</td>  
                        ${usuarioAnalista() ? `  
                            <td class="text-center">  
                                <button type="button" class="btn btn-sm btn-danger" onclick="removerNota(${index})" title="Remover">  
                                    <i class="bi bi-trash"></i>  
                                </button>  
                            </td>  
                        ` : ''}  
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

            calcularCoeficienteNQA();
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

        function abrirModalObservacao() {
            $('#observacaoDescricao').val('');
            const modal = new bootstrap.Modal(document.getElementById('modalObservacao'));
            modal.show();
        }

        function salvarObservacao() {
            const descricao = $('#observacaoDescricao').val().trim();

            if (!descricao) {
                Toast.warning('Informe a observação.');
                return;
            }

            listaObservacoes.push({
                id: null,
                descricao: descricao
            });

            Toast.success('Observação adicionada com sucesso!');
            renderizarObservacoes();
            bootstrap.Modal.getInstance(document.getElementById('modalObservacao')).hide();
        }

        function removerObservacao(index) {
            ConfirmationModal.danger({
                message: 'Deseja realmente remover essa observação?',
                type: 'danger',
                confirmText: 'Remover',
                cancelText: 'Cancelar',
                onConfirm: () => {
                    listaObservacoes.splice(index, 1);
                    renderizarObservacoes();
                    Toast.success('Observação removida com sucesso!');
                },
                onCancel: () => {}
            });
        }

        function renderizarObservacoes() {
            const $container = $('#observacoesListaContainer');
            $('#contadorObservacoes, #menuContadorObservacoes').text(listaObservacoes.length);

            if (listaObservacoes.length === 0) {
                $container.html(`
                    <div class="text-center py-5">
                        <p class="text-muted mt-3 mb-0">Nenhuma observação adicionada.</p>
                    </div>
                `);
                return;
            }

            let html = `
                <div class="table-responsive">
                    <table class="table table-hover table-sm mb-0">
                        <thead class="table-light">
                            <tr>
                                <th>Data Criação</th>
                                <th>Usuário</th>
                                <th>Descrição</th>
                                <th class="text-center" style="width: 80px;">Ações</th>
                            </tr>
                        </thead>
                        <tbody>
            `;

            listaObservacoes.forEach((observacao, index) => {
                html += `
                    <tr>
                        <td>${observacao.dataCriacao}</td>
                        <td>${observacao.nomeUsuarioInclusao}</td>
                        ${criarCelulaDetalhe(observacao.descricao, 'Observação')}
                        <td class="text-center">
                            <button type="button" class="btn btn-sm btn-danger" onclick="removerObservacao(${index})" title="Remover">
                                <i class="bi bi-trash"></i>
                            </button>
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
        }

        function renderizarRespostas() {
            $('#contadorRespostas, #menuContadorRespostas').text(listaRespostas.length);

            let html = `
                <div style="overflow: auto; max-height: 45vh;" class="table-responsive">
                    <table class="table table-sm table-hover mb-0">
                        <thead class="table-light">
                            <tr>
                                <th>Data Envio</th>
                                <th>Data Resposta</th>
                                <th>Usuário Envio</th>
                                <th>Usuário Resposta</th>
                                <th>Procede</th>
                                <th>Descrição</th>
                                <th>Justificativa</th>
                            </tr>
                        </thead>
                        <tbody>
            `;

            listaRespostas.forEach(item => {
                html += `
                    <tr>
                        <td>${item.dataInclusao || '-'}</td>
                        <td>${item.dataFinalizacao || '-'}</td>
                        <td>${item.nomeUsuarioInclusao || '-'}</td>
                        <td>${item.nomeUsuarioFinalizacao || '-'}</td>
                        <td>${item.avaliacao == '' ? 'Ainda não avaliado' : item.avaliacao == 1 ? 'Sim' : 'Não'}</td>
                        ${criarCelulaDetalhe(item.descricao, 'Descrição da resposta')}
                        ${criarCelulaDetalhe(item.justificativa, 'Justificativa da resposta')}
                    </tr>
                `;
            });

            html += `
                        </tbody>
                    </table>
                </div>
            `;

            $('#respostasListaContainer').html(html);
        }

        function renderizarAprovacoes() {
            $('#contadorAprovacoes, #menuContadorAprovacoes').text(listaAprovacoes.length);

            let html = `
                <div style="overflow: auto; max-height: 45vh;" class="table-responsive">
                    <table class="table table-sm table-hover mb-0">
                        <thead class="table-light">
                            <tr>
                                <th>Data Envio</th>
                                <th>Data Avaliação</th>
                                <th>Usuário Envio</th>
                                <th>Usuário Avaliação</th>
                                <th>Nível</th>
                                <th>Aprovado</th>
                                <th>Justificativa</th>
                            </tr>
                        </thead>
                        <tbody>
            `;

            listaAprovacoes.forEach(item => {
                html += `
                    <tr style="${item.invalido == '1' ? 'background-color: #f9e4e4;' : ''}">
                        <td>${item.dataInclusao || '-'}</td>
                        <td>${item.dataAvaliacao || '-'}</td>
                        <td>${item.nomeUsuarioInclusao || '-'}</td>
                        <td>${item.nomeUsuarioAvaliacao || '-'}</td>
                        <td>${item.nivel || '-'}</td>
                        <td>${item.avaliacao || '-'}</td>
                        ${criarCelulaDetalhe(item.motivo, 'Justificativa da aprovação')}
                    </tr>
                `;
            });

            html += `
                        </tbody>
                    </table>
                </div>
            `;

            $('#aprovacoesListaContainer').html(html);
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
                        <div class="d-flex gap-2">
                            <a href="/Clientes_Docs/${anexo.caminho}" target="_blank" class="btn btn-sm btn-primary">
                                <i class="bi bi-box-arrow-up-right me-1"></i>Abrir
                            </a>
                            ${usuarioAnalista() ? `
                                <button type="button" class="btn btn-sm btn-danger" onclick="removerAnexo(${index})" title="Remover">
                                    <i class="bi bi-trash"></i>
                                </button>
                            ` : ''}
                        </div>
                    </div>
                `;
            });

            html += '</div>';
            $container.html(html);
        }

        function removerAnexo(index) {
            ConfirmationModal.danger({
                message: 'Deseja realmente remover este anexo?',
                type: 'danger',
                confirmText: 'Remover',
                cancelText: 'Cancelar',
                onConfirm: () => {
                    listaAnexos.splice(index, 1);
                    renderizarAnexos();
                    Toast.success('Anexo removido com sucesso!');
                },
                onCancel: () => {}
            });
        }

        function renderizarIshikawa(causasIshikawa, descricaoProblema) {
            ishikawa = new Ishikawa({
                el: '#ishikawa-container',
                problema: descricaoProblema,
                readOnly: true
            });

            ishikawa.setData((causasIshikawa || []).map(causa => {
                return {
                    ...causa,
                    causaRaiz: causa.causaRaiz == '1'
                };
            }));
        }

        function renderizar5W2H(acoes5W2H, descricaoProblema) {
            W2H = new W2H({
                el: '#5W2H-container',
                assunto: descricaoProblema,
                readOnly: true
            });

            W2H.setData(acoes5W2H);
        }

        function inicializarValidacoes() {
            const naoVazio = (msg) => ({
                validator: (value) => value !== null && value !== undefined && String(value).trim() !== '',
                errorMessage: msg
            });

            const dataValida = (msg) => ({
                validator: (value) => {
                    const regex = /^\d{4}-(0[1-9]|1[0-2])-(0[1-9]|[12]\d|3[01])$/;
                    if (!regex.test(value)) return false;

                    const [year, month, day] = value.split('-').map(Number);
                    const date = new Date(year, month - 1, day);

                    return date.getFullYear() === year &&
                        date.getMonth() === month - 1 &&
                        date.getDate() === day;
                },
                errorMessage: msg
            });

            validacoes.aba1 = new JustValidate('#formularioFNC', { lockForm: false, validateBeforeSubmitting: true });
            validacoes.aba1
                .addField('#codPlanta', [naoVazio('Planta é obrigatória')])
                .addField('#codCliente', [naoVazio('Cliente é obrigatório')])
                .addField('#codLoja', [naoVazio('Loja é obrigatória')])
                .addField('#dataOcorrencia', [naoVazio('Data de ocorrência é obrigatória'), dataValida('Data de ocorrência inválida')])
                .addField('#tipoDefeito', [naoVazio('Tipo de defeito é obrigatório')])
                .addField('#defeitoId', [naoVazio('Classificação de defeito é obrigatória')])
                .addField('#disposicaoId', [naoVazio('Código disposição é obrigatório')])
                .addField('#quantidadeReclamada', [naoVazio('Quantidade reclamada é obrigatória')])
                .addField('#valorReclamado', [naoVazio('Valor reclamado é obrigatório')]);

            validacoes.aba2 = new JustValidate('#formularioFNC', { lockForm: false });

            validacoes.aba3 = new JustValidate('#formularioFNC', { lockForm: false });
            validacoes.aba3
                .addField('#status', [naoVazio('Status é obrigatório')])
                .addField('#situacao', [naoVazio('Situação é obrigatória')])
                .addField('#setorId', [naoVazio('Setor é obrigatório')]);

            validacoes.aba4 = new JustValidate('#formularioFNC', { lockForm: false });
            validacoes.aba5 = new JustValidate('#formularioFNC', { lockForm: false });
            validacoes.aba6 = new JustValidate('#formularioFNC', { lockForm: false });
            validacoes.aba7 = new JustValidate('#formularioFNC', { lockForm: false });
            validacoes.aba8 = new JustValidate('#formularioFNC', { lockForm: false });
            validacoes.aba9 = new JustValidate('#formularioFNC', { lockForm: false });
            validacoes.aba10 = new JustValidate('#formularioFNC', { lockForm: false });
        }

        function navegarParaGrade() {
            window.location.href = 'FNCGrade.asp';
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

        function usuarioAnalista() {
            return String(nivelUsuario || '').trim().toUpperCase() === 'ANALISTA';
        }

        function bloquearFormularioSomenteAnexo() {
            formularioSomenteAnexo = !usuarioAnalista();

            if (!formularioSomenteAnexo) {
                return;
            }

            $('#formularioFNC')
                .find('input, select, textarea, button')
                .not('#uploadAnexos input, #uploadAnexos select, #uploadAnexos textarea, #uploadAnexos button')
                .prop('disabled', true);

            $('.select2').prop('disabled', true).trigger('change.select2');
            $('button[onclick="abrirModalNota()"]').addClass('d-none');
            $('button[onclick="abrirModalObservacao()"]').addClass('d-none');
            $('#btnEnviarParaResponsavel, #btnEnviarParaAprovacao').addClass('d-none').removeClass('d-inline-flex');
            $('#divAcoes').removeClass('d-none').addClass('d-flex');
            $('#btnSalvar').prop('disabled', false).removeClass('d-none');
        }

        function atualizarBotoesAcao(dados) {
            $('#divAcoes').removeClass('d-none').addClass('d-flex');

            if (!usuarioAnalista()) {
                $('#btnEnviarParaResponsavel, #btnEnviarParaAprovacao').addClass('d-none').removeClass('d-inline-flex');
                return;
            }

            const podeEnviarParaResponsavel = dados.situacao != eSituacao.ENCERRADA;
            const podeEnviarParaAprovacao = 1;

            if (podeEnviarParaResponsavel) {
                $('#btnEnviarParaResponsavel').removeClass('d-none').addClass('d-inline-flex');
            }

            if (podeEnviarParaAprovacao) {
                $('#btnEnviarParaAprovacao').removeClass('d-none').addClass('d-inline-flex');
            }
        }

        function marcarAbaComoCompleta(numeroAba) {
            const $button = $(`.tab-button[data-tab="${numeroAba}"]`);

            if ($button.length) {
                $button.addClass('completed').removeClass('invalid');
                $button.find('.tab-icon').attr('class', 'tab-icon bi bi-check-circle-fill');
            }

            abasValidadas.add(numeroAba);
        }

        function marcarAbaComoInvalida(numeroAba) {
            const $button = $(`.tab-button[data-tab="${numeroAba}"]`);

            if ($button.length && !$button.hasClass('completed')) {
                $button.addClass('invalid');
                $button.find('.tab-icon').attr('class', 'tab-icon bi bi-exclamation-circle-fill');
            }
        }

        function carregarDadosEdicao() {
            $.ajax({
                url: 'FNCRepositorio.asp',
                type: 'GET',
                data: {
                    acao: 'obterPorId',
                    id: id
                },
                dataType: 'json',
                success: function(response) {
                    if (!response.sucesso) {
                        Toast.error(response.mensagem);
                        return;
                    }

                    preencherFormulario(response.dados);
                    atualizarBotoesAcao(response.dados);
                },
                error: function(xhr, status, error) {
                    console.error('Erro ao carregar dados para edição:', error);
                    Toast.error('Erro ao carregar dados para edição');
                }
            });
        }

        function preencherFormulario(dados) {
            var p1 = carregarDefeitos(dados.tipoDefeito);
            var p2 = carregarResponsaveis(dados.codPlanta);
            var p3 = carregarDisposicoes();
            var p4 = carregarTiposDefeito();
            var p5 = carregarStatus();
            var p6 = carregarSituacoes();
            var p7 = carregarSetores();

            $.when(p1, p2, p3, p4, p5, p6, p7).then(function() {
                $('#btnBuscarCliente').prop('disabled', true);
                $('#nomeCliente').prop('disabled', true).val(dados.nomeCliente);
                $('#codCliente').prop('disabled', true).val(dados.codCliente);
                $('#codLoja').prop('disabled', true).val(dados.codLoja);
                $('#quantidadeReclamada').val(dados.quantidadeReclamada).trigger('input');
                $('#valorReclamado').val(dados.valorReclamado).maskMoney('mask', dados.valorReclamado);
                $('#descricaoDisposicao').val(dados.descricaoDisposicao);
                $('#descricaoDefeito').val(dados.descricaoDefeito);
                $('#contato').val(dados.contato);
                $('#email').val(dados.email);
                $('#telefone').val(dados.telefone);
                $('#dataOcorrencia').val(dados.dataOcorrencia);
                $('#dataConclusaoPrevista').val(dados.dataConclusaoPrevista);
                $('#dataEnvioCliente').val(dados.dataEnvioCliente);
                $('#dataProducao').val(dados.dataProducao);
                $('#turno').val(dados.turno);
                $('#maquina').val(dados.maquina);
                $('#eficacia').val(dados.eficacia);
                $('#cidadeEstado').val(dados.cidadeEstado);

                setTimeout(() => {
                    $('#codPlanta').prop('disabled', true).val(dados.codPlanta).trigger('change.select2');
                    $('#setorId').val(dados.setorId).trigger('change.select2');
                    $('#defeitoId').val(dados.defeitoId).trigger('change.select2');
                    $('#responsavelId').val(dados.responsavelId).trigger('change.select2');
                    $('#tipoDefeito').val(dados.tipoDefeito).trigger('change.select2');
                    $('#disposicaoId').val(dados.disposicaoId).trigger('change.select2');
                    $('#status').val(dados.status).trigger('change.select2');
                    $('#situacao').val(dados.situacao).trigger('change.select2');

                    calcularCoeficienteNQA();
                    $('.tab-button').prop('disabled', false);
                    bloquearFormularioSomenteAnexo();
                }, 500);

                listaNotas = dados.notas || [];
                listaAnexos = dados.anexos || [];
                listaObservacoes = dados.observacoes || [];
                listaRespostas = dados.respostas || [];
                listaAprovacoes = dados.aprovacoes || [];

                renderizarNotas();
                renderizarAnexos();
                renderizarObservacoes();
                renderizarAprovacoes();
                renderizarRespostas();
                renderizarIshikawa(dados.causasIshikawa, dados.defeitoIshikawa || dados.descricaoTipoDefeito);
                renderizar5W2H(dados.acoes5W2H, dados.assunto5W2H || dados.descricaoTipoDefeito);
            });
        }

        function ehTipoDefeitoIndustrial() {
            const $tipoDefeito = $('#tipoDefeito');
            const valor = String($tipoDefeito.val() || '').trim().toUpperCase();
            const texto = String($tipoDefeito.find('option:selected').text() || '').trim().toUpperCase();

            return valor === 'INDUSTRIAL' || texto === 'INDUSTRIAL';
        }

        function calcularCoeficienteNQA() {
            const box = document.getElementById('nqa-inline');

            if (!ehTipoDefeitoIndustrial() || !box) {
                if (box) box.classList.remove('visible', 's-ok', 's-warn', 's-fail');
                return;
            }

            const totalQuantidadeNotas = listaNotas.reduce((total, nota) => total + parseFloat(String(nota.quantidade).replace(',', '.')), 0);
            const quantidadeReclamada = parseFloat($('#quantidadeReclamada').val().replace(/\D/g, ''));

            if (isNaN(totalQuantidadeNotas) || totalQuantidadeNotas < 1 || isNaN(quantidadeReclamada) || quantidadeReclamada < 0 || quantidadeReclamada > totalQuantidadeNotas) {
                box.classList.remove('visible', 's-ok', 's-warn', 's-fail');
                return;
            }

            const nqa = (quantidadeReclamada / totalQuantidadeNotas) * 100;
            let state, badge, detail;

            if(parseFloat(nqa) < parseFloat(limitesNQA.critico)) {
                state = 's-ok';
                badge = 'Aprovado';
                detail = `Menor do que ${limitesNQA.critico}%`;
            } else if ( (parseFloat(nqa) >= parseFloat(limitesNQA.critico)) && (parseFloat(nqa) < parseFloat(limitesNQA.grave)) ) {
                state = 's-fail';
                badge = 'Crítico';
                detail = `Maior ou igual a ${limitesNQA.critico}%`;
            } else {
                state = 's-fail';
                badge = 'Grave';
                detail = `Maior ou igual a ${limitesNQA.grave}%`;
            }

            $('#nq-val').text(formatarMoeda(nqa) + '%');
            $('#nq-badge').text(badge);
            $('#nq-detail').text(detail);

            box.classList.remove('s-ok', 's-warn', 's-fail');
            box.classList.add('visible', state);
        }

        function anexarEventos() {
            $('#tipoDefeito').on('change', function() {
                carregarDefeitos($(this).val());
                calcularCoeficienteNQA();
            });

            $('#formularioVinculoNota').on('submit', function(e) {
                e.preventDefault();
                pesquisarNotas();
            });

            $('#btnSalvarNota').on('click', salvarNota);
            $('#btnSalvarObservacao').on('click', salvarObservacao);

            $('#btnSalvar').on('click', async function() {
                const p1 = await validarFormulario();

                if (p1 && p1.sucesso) {
                    Toast.success('FNC salva com sucesso!');
                    setTimeout(() => {
                        navegarParaGrade();
                    }, 1000);
                }
            });

            $('#btnEnviarParaResponsavel').on('click', async function() {
                var p1 = await validarFormulario();
                if (!p1 || !p1.sucesso) return;

                var p2 = await enviarParaResponsavel();
                if (p2.sucesso) {
                    Toast.success('FNC enviada com sucesso!');
                    setTimeout(() => {
                        navegarParaGrade();
                    }, 1000);
                }
            });

            $('#btnEnviarParaAprovacao').on('click', async function() {
                var p1 = await validarFormulario();
                if (!p1 || !p1.sucesso) return;

                var p2 = await enviarParaAprovacao();
                if (p2.sucesso) {
                    Toast.success('FNC enviada com sucesso!');
                    setTimeout(() => {
                        navegarParaGrade();
                    }, 1000);
                }
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

        function irProximo(){
            navegarParaAba(abaAtual + 1);
        }

        function irAnterior(){
            navegarParaAba(abaAtual - 1);
        }

        async function validarFormulario() {
            if (formularioSomenteAnexo) {
                return enviarFormulario();
            }

            const abas = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10];

            const resultados = await Promise.all([
                validacoes.aba1.revalidate(),
                Promise.resolve(listaNotas.length > 0 && notasComQuantidadeReclamada()),
                validacoes.aba3.revalidate(),
                validacoes.aba4.revalidate(),
                validacoes.aba5.revalidate(),
                validacoes.aba6.revalidate(),
                validacoes.aba7.revalidate(),
                validacoes.aba8.revalidate(),
                validacoes.aba9.revalidate(),
                validacoes.aba10.revalidate()
            ]);

            resultados.forEach((valida, index) => {
                const numeroAba = abas[index];

                if (valida) {
                    marcarAbaComoCompleta(numeroAba);
                } else {
                    marcarAbaComoInvalida(numeroAba);
                    exibirErrosValidacao(validacoes[`aba${numeroAba}`]);
                }
            });

            if (listaNotas.length == 0) {
                Toast.error('Adicione ao menos uma nota.');
            } else if (!notasComQuantidadeReclamada()) {
                Toast.error('Informe a quantidade reclamada de todas as notas');
            }

            if (!resultados.every(Boolean)) {
                const primeiraInvalida = abas[resultados.findIndex(v => !v)];
                navegarParaAba(primeiraInvalida);
                Toast.error('Existem campos obrigatórios pendentes. Verifique as seções destacadas.');
                return { sucesso: false };
            }

            return enviarFormulario();
        }

        function exibirLoading(exibir) {
            if (exibir) {
                $('#loadingOverlay').removeClass('d-none').addClass('d-flex');
            } else {
                $('#loadingOverlay').removeClass('d-flex').addClass('d-none');
            }
        }

        function enviarParaResponsavel() {
            exibirLoading(true);

            return $.ajax({
                url: 'FNCRepositorio.asp',
                type: 'POST',
                data: {
                    acao: 'enviarParaResponsavel',
                    idFNC: id,
                    responsavelId: $('#responsavelId').val()
                },
                dataType: 'json',
                success: function(response) {
                    if (!response.sucesso) {
                        Toast.error(response.mensagem);
                    }
                },
                error: function() {
                    Toast.error('Erro ao enviar para responsável');
                },
                complete: function() {
                    exibirLoading(false);
                }
            });
        }

        function enviarParaAprovacao() {
            exibirLoading(true);

            return $.ajax({
                url: 'FNCRepositorio.asp',
                type: 'POST',
                data: {
                    acao: 'enviarParaAprovacao',
                    idFNC: id
                },
                dataType: 'json',
                success: function(response) {
                    if (!response.sucesso) {
                        Toast.error(response.mensagem);
                    }
                },
                error: function() {
                    Toast.error('Erro ao enviar para aprovação');
                },
                complete: function() {
                    exibirLoading(false);
                }
            });
        }

        function enviarFormulario() {
            exibirLoading(true);

            const acoes5W2H = W2H ? W2H.getData() : [];
            const causasIshikawa = ishikawa ? ishikawa.getData().map(causa => {
                return {
                    ...causa,
                    causaRaiz: causa.causaRaiz ? 1 : 0
                };
            }) : [];

            var novosAnexos = typeof uploadAnexos !== 'undefined' && uploadAnexos ? uploadAnexos.getFiles() : [];
            novosAnexos = novosAnexos.map(anexo => {
                return {
                    ...anexo,
                    id: ''
                };
            });

            const anexosExistentes = listaAnexos.map(anexo => {
                return {
                    id: anexo.id,
                    name: anexo.nome,
                    uploadedName: anexo.caminho
                };
            });

            const anexos = [...anexosExistentes, ...novosAnexos];
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

            const dadosFormulario = {
                acao: 'editar',
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
                dataEnvioCliente: $('#dataEnvioCliente').val(),
                dataConclusaoPrevista: $('#dataConclusaoPrevista').val(),
                dataProducao: $('#dataProducao').val(),
                status: $('#status').val(),
                situacao: $('#situacao').val(),
                eficacia: $('#eficacia').val(),
                maquina: $('#maquina').val(),
                turno: $('#turno').val(),
                setorId: $('#setorId').val(),
                responsavelId: $('#responsavelId').val(),
                defeitoIshikawa: ishikawa ? ishikawa.opts.problema : '',
                causasIshikawa: JSON.stringify(causasIshikawa),
                notas: JSON.stringify(notas),
                anexos: JSON.stringify(anexos),
                observacoes: JSON.stringify(listaObservacoes),
                assunto5W2H: W2H ? W2H.opts.assunto : '',
                acoes5W2H: JSON.stringify(acoes5W2H)
            };

            if (id) {
                dadosFormulario.id = id;
            }

            return $.ajax({
                url: 'FNCRepositorio.asp',
                type: 'POST',
                data: dadosFormulario,
                dataType: 'json',
                success: function(response) {
                    if (!response.sucesso) {
                        Toast.error(response.mensagem);
                    }
                },
                error: function() {
                    Toast.error('Erro ao salvar formulário');
                },
                complete: function() {
                    exibirLoading(false);
                }
            });
        }

        function obterNivelUsuario() {
            return new Promise((resolve, reject) => {
                $.ajax({
                    url: 'FNCRepositorio.asp?v=<%=Now()%>',
                    type: 'GET',
                    data: {
                        acao: 'obterNivelUsuario',
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

        $(document).ready(async function() {
            resposta = await obterNivelUsuario();
            if (!resposta.sucesso) {
                alert("Você não tem permissão para acessar esta página.");
                window.location.href = 'index.asp';
            }
        
            nivelUsuario = resposta.dados?.nivel;
            if (!nivelUsuario) {
                alert("Você não tem permissão para acessar esta página.");
                window.location.href = 'index.asp';
            }

            $('.tab-button').prop('disabled', true);

            const urlParams = new URLSearchParams(window.location.search);
            const paramId = urlParams.get('id');

            if (paramId) {
                id = paramId;
                $('#idTitulo').text(id);
            }

            carregarDadosEdicao();
            carregarLimitesNQA();
            anexarEventos();
            navegarParaAba(1);

            inicializarValidacoes();
            inicializarSelect2();
            inicializarMascaras();
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
                <h3 class="mb-0">Editar FNC #<span id="idTitulo"></span></h3>
            </div>

            <div id="divAcoes" class="d-none page-actions flex-wrap">
                <button type="button" class="btn btn-sm btn-success" id="btnSalvar">
                    <i class="bi bi-check-lg me-1"></i>
                    Salvar alterações
                </button>

                <button type="button" class="d-none btn btn-sm btn-outline-secondary" id="btnEnviarParaResponsavel">
                    <i class="bi bi-person-check me-1"></i>
                    Salvar e enviar para Responsável
                </button>

                <button type="button" class="d-none btn btn-sm btn-outline-primary" id="btnEnviarParaAprovacao">
                    <i class="bi bi-shield-check me-1"></i>
                    Salvar e enviar para Aprovação
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
                            <i class="tab-icon bi bi-search"></i>
                            <span class="tab-title">Análise</span>
                        </button>

                        <button type="button" class="tab-button" data-tab="4">
                            <i class="tab-icon bi bi-person-lines-fill"></i>
                            <span class="tab-title">Contato</span>
                        </button>

                        <div class="tab-group-title">Ferramentas de Qualidade</div>

                        <button type="button" class="tab-button" data-tab="5">
                            <i class="tab-icon bi bi-diagram-3"></i>
                            <span class="tab-title">Ishikawa</span>
                        </button>

                        <button type="button" class="tab-button" data-tab="6">
                            <i class="tab-icon bi bi-list-check"></i>
                            <span class="tab-title">Plano 5W2H</span>
                        </button>

                        <div class="tab-group-title">Adicionais</div>

                        <button type="button" class="tab-button" data-tab="7">
                            <i class="tab-icon bi bi-paperclip"></i>
                            <span class="tab-title">Anexos</span>
                            <span class="tab-meta" id="menuContadorAnexos">0</span>
                        </button>

                        <button type="button" class="tab-button" data-tab="8">
                            <i class="tab-icon bi bi-chat-dots"></i>
                            <span class="tab-title">Observações</span>
                            <span class="tab-meta" id="menuContadorObservacoes">0</span>
                        </button>

                        <div class="tab-group-title">Histórico</div>

                        <button type="button" class="tab-button" data-tab="9">
                            <i class="tab-icon bi bi-reply"></i>
                            <span class="tab-title">Respostas</span>
                            <span class="tab-meta" id="menuContadorRespostas">0</span>
                        </button>

                        <button type="button" class="tab-button" data-tab="10">
                            <i class="tab-icon bi bi-shield-check"></i>
                            <span class="tab-title">Aprovações</span>
                            <span class="tab-meta" id="menuContadorAprovacoes">0</span>
                        </button>
                    </div>
                </div>

                <div class="tab-content-wrapper">
                    <form id="formularioFNC" name="formularioFNC" novalidate>
                        <div class="tab-content active" data-tab-content="1">
                            <div class="section-header">
                                <div class="section-title">Cadastro da reclamação</div>
                                <p class="section-subtitle">
                                    Dados principais da FNC
                                </p>
                            </div>

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

                                <div class="col-12 col-md-2">
                                    <label class="form-label">Buscar cliente</label>
                                    <div class="input-group">
                                        <button id="btnBuscarCliente" type="button" class="flex-grow-1 btn btn-sm btn-secondary">
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

                                <div class="col-12 col-md-3">
                                    <label for="dataOcorrencia" class="form-label">Data de Ocorrência <span class="required-field">*</span></label>
                                    <input type="date" class="form-control form-control-sm" id="dataOcorrencia" name="dataOcorrencia">
                                </div>

                                <div class="col-12 col-md-3">
                                    <label for="tipoDefeito" class="form-label">Tipo de Defeito <span class="required-field">*</span></label>
                                    <select class="form-select form-select-sm select2" name="tipoDefeito" id="tipoDefeito">
                                        <option value="">Selecione o tipo de defeito</option>
                                    </select>
                                </div>

                                <div class="col-12 col-md-3">
                                    <label for="defeitoId" class="form-label">Classificação Defeito <span class="required-field">*</span></label>
                                    <select class="form-select form-select-sm select2" name="defeitoId" id="defeitoId">
                                        <option value="">Selecione um defeito</option>
                                    </select>
                                </div>

                                <div class="col-12 col-md-3">
                                    <label for="disposicaoId" class="form-label">Código Disposição <span class="required-field">*</span></label>
                                    <select class="form-select form-select-sm select2" name="disposicaoId" id="disposicaoId">
                                        <option value="">Selecione a disposição</option>
                                    </select>
                                </div>

                                <div class="col-12">
                                    <label for="descricaoDisposicao" class="form-label">Disposição <span class="required-field">*</span></label>
                                    <textarea disabled class="form-control form-control-sm" id="descricaoDisposicao" name="descricaoDisposicao" rows="4"></textarea>
                                </div>

                                <div class="col-12">
                                    <label for="descricaoDefeito" class="form-label">Descrição Defeito <span class="required-field">*</span></label>
                                    <textarea disabled class="form-control form-control-sm" id="descricaoDefeito" name="descricaoDefeito" rows="4"></textarea>
                                </div>
                            </div>
                        </div>

                        <div class="tab-content" data-tab-content="2">
                            <div class="section-header d-flex justify-content-between align-items-center">
                                <div>
                                    <div class="section-title">
                                        Notas vinculadas <span class="badge bg-secondary ms-1" id="contadorNotas">0</span>
                                    </div>
                                    <p class="section-subtitle">Notas fiscais relacionadas à reclamação.</p>
                                </div>

                                <button type="button" class="btn btn-sm btn-primary" onclick="abrirModalNota()">
                                    <i class="bi bi-plus-lg me-1"></i>Adicionar
                                </button>
                            </div>

                            <div class="row g-3 mb-3">
                                <div class="col-lg-4 col-md-6">
                                    <label for="quantidadeReclamada" class="form-label">Quantidade Reclamada Total</label>
                                    <input onblur="calcularCoeficienteNQA()" type="text" class="form-control form-control-sm" id="quantidadeReclamada" value="0" readonly>
                                </div>
                                <div class="col-lg-4 col-md-6">
                                    <label for="valorReclamado" class="form-label">Valor Reclamado Total</label>
                                    <input onblur="calcularCoeficienteNQA()" type="text" class="form-control form-control-sm" id="valorReclamado" value="0,00" >
                                </div>
                            </div>
                            <div id="notasListaContainer"></div>
                        </div>
 
                        <div class="tab-content" data-tab-content="3">
                            <div class="section-header">
                                <div class="section-title">Análise interna</div>
                                <p class="section-subtitle">Status, situação, setor responsável, NQA e informações produtivas.</p>
                            </div>

                            <div class="nqa-inline" id="nqa-inline">
                                <span class="nq-row-label">Quociente NQA:</span>
                                <span class="nq-val" id="nq-val">—</span>
                                <span class="nq-sep"></span>
                                <span class="nq-badge" id="nq-badge">—</span>
                                <span class="nq-detail" id="nq-detail"></span>
                            </div>

                            <div class="row g-3">
                                <div class="col-lg-4 col-md-6">
                                    <label for="dataEnvioCliente" class="form-label">Data de Envio Cliente</label>
                                    <input type="date" class="form-control form-control-sm" id="dataEnvioCliente" name="dataEnvioCliente">
                                </div>

                                <div class="col-lg-4 col-md-6">
                                    <label for="dataConclusaoPrevista" class="form-label">Data de Conclusão Prevista / Ponto de Corte</label>
                                    <input type="date" class="form-control form-control-sm" id="dataConclusaoPrevista" name="dataConclusaoPrevista">
                                </div>

                                <div class="col-lg-4 col-md-6">
                                    <label for="dataProducao" class="form-label">Data de Produção</label>
                                    <input type="date" class="form-control form-control-sm" id="dataProducao" name="dataProducao">
                                </div>

                                <div class="col-lg-4 col-md-6">
                                    <label for="status" class="form-label">Status <span class="required-field">*</span></label>
                                    <select class="form-select form-select-sm select2" name="status" id="status">
                                        <option value="">Selecione um status</option>
                                    </select>
                                </div>

                                <div class="col-lg-4 col-md-6">
                                    <label for="situacao" class="form-label">Situação <span class="required-field">*</span></label>
                                    <select class="form-select form-select-sm select2" name="situacao" id="situacao">
                                        <option value="">Selecione uma situação</option>
                                    </select>
                                </div>

                                <div class="col-lg-4 col-md-6">
                                    <label for="setorId" class="form-label">Setor <span class="required-field">*</span></label>
                                    <select class="form-select form-select-sm select2" name="setorId" id="setorId">
                                        <option value="">Selecione um setor</option>
                                    </select>
                                </div>

                                <div class="col-lg-4 col-md-6">
                                    <label for="responsavelId" class="form-label">Responsável</label>
                                    <select class="form-select form-select-sm select2" name="responsavelId" id="responsavelId">
                                        <option value="">Selecione um responsável</option>
                                    </select>
                                </div>

                                <div class="col-lg-4 col-md-6">
                                    <label for="turno" class="form-label">Turno</label>
                                    <input type="text" class="form-control form-control-sm" id="turno" name="turno">
                                </div>

                                <div class="col-lg-4 col-md-6">
                                    <label for="maquina" class="form-label">Máquina</label>
                                    <input type="text" class="form-control form-control-sm" id="maquina" name="maquina">
                                </div>

                                <div class="col-12">
                                    <label for="eficacia" class="form-label">Eficácia/Observação</label>
                                    <textarea class="form-control form-control-sm" id="eficacia" name="eficacia" rows="4"></textarea>
                                </div>
                            </div>
                        </div>

                        <div class="tab-content" data-tab-content="4">
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

                        <div class="tab-content" data-tab-content="5">
                            <div class="section-header">
                                <div class="section-title">Diagrama de Ishikawa</div>
                                <p class="section-subtitle">Análise de causa raiz do problema.</p>
                            </div>

                            <div id="ishikawa-container"></div>
                        </div>

                        <div class="tab-content" data-tab-content="6">
                            <div class="section-header">
                                <div class="section-title">Plano de ação 5W2H</div>
                                <p class="section-subtitle">Ações, responsáveis, prazos e tratativas da não conformidade.</p>
                            </div>

                            <div id="5W2H-container"></div>
                        </div>

                        <div class="tab-content" data-tab-content="7">
                            <div class="section-header">
                                <div class="section-title">Anexos</div>
                                <p class="section-subtitle">Evidências, documentos e arquivos relacionados à FNC.</p>
                            </div>

                            <div class="row g-3">
                                <file-upload id="uploadAnexos" file-code="FNC" max-files="10" multiple></file-upload>
                                <div id="anexosListaContainer"></div>
                            </div>
                        </div>

                        <div class="tab-content" data-tab-content="8">
                            <div class="section-header d-flex justify-content-between align-items-center">
                                <div>
                                    <div class="section-title">
                                        Observações <span class="badge bg-secondary ms-1" id="contadorObservacoes">0</span>
                                    </div>
                                    <p class="section-subtitle">Comentários e registros adicionais da FNC.</p>
                                </div>

                                <button type="button" class="btn btn-sm btn-primary" onclick="abrirModalObservacao()">
                                    <i class="bi bi-plus-lg me-1"></i>Adicionar
                                </button>
                            </div>

                            <div id="observacoesListaContainer"></div>
                        </div>

                        <div class="tab-content" data-tab-content="9">
                            <div class="section-header">
                                <div class="section-title">
                                    Histórico de respostas <span class="badge bg-secondary ms-1" id="contadorRespostas">0</span>
                                </div>
                                <p class="section-subtitle">Consulta dos retornos registrados para esta FNC.</p>
                            </div>

                            <div id="respostasListaContainer"></div>
                        </div>

                        <div class="tab-content" data-tab-content="10">
                            <div class="section-header">
                                <div class="section-title">
                                    Histórico de aprovações <span class="badge bg-secondary ms-1" id="contadorAprovacoes">0</span>
                                </div>
                                <p class="section-subtitle">Consulta das avaliações e aprovações registradas.</p>
                            </div>

                            <div id="aprovacoesListaContainer"></div>
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

                    <div id="resultadosPesquisaNota"></div>
                </div>
                <div class="modal-footer">
                    <button type="button" class="btn btn-sm btn-secondary" data-bs-dismiss="modal">Cancelar</button>
                    <button type="button" class="btn btn-sm btn-primary" id="btnSalvarNota" disabled>
                        <i class="bi bi-plus-lg me-1"></i>Adicionar
                    </button>
                </div>
            </div>
        </div>
    </div>

    <div class="modal fade" id="modalObservacao" tabindex="-1" aria-labelledby="modalObservacaoLabel" aria-hidden="true">
        <div class="modal-dialog modal-xl">
            <div class="modal-content">
                <div class="modal-header">
                    <h5 class="modal-title" id="modalObservacaoTitulo">
                        <i class="bi bi-chat-dots me-2"></i>Adicionar Observação
                    </h5>
                    <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Fechar"></button>
                </div>
                <div class="modal-body">
                    <form id="formularioObservacao" name="formularioObservacao" novalidate>
                        <div class="col-12">
                            <label for="observacaoDescricao" class="form-label">Observação</label>
                            <textarea required class="form-control form-control-sm" id="observacaoDescricao" name="observacaoDescricao" rows="4"></textarea>
                        </div>
                    </form>
                </div>
                <div class="modal-footer">
                    <button type="button" class="btn btn-sm btn-secondary" data-bs-dismiss="modal">Cancelar</button>
                    <button type="button" class="btn btn-sm btn-primary" id="btnSalvarObservacao">
                        <i class="bi bi-plus-lg me-1"></i>Adicionar
                    </button>
                </div>
            </div>
        </div>
    </div>
    
    <div class="modal fade" id="modalDetalheTexto" tabindex="-1" aria-labelledby="modalDetalheTituloTexto" aria-hidden="true">
        <div class="modal-dialog modal-lg modal-dialog-scrollable">
            <div class="modal-content">
                <div class="modal-header">
                    <h5 class="modal-title" id="modalDetalheTituloTexto">
                        Detalhes
                    </h5>
                    <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Fechar"></button>
                </div>
                <div class="modal-body">
                    <div id="modalDetalheConteudo" class="detalhe-conteudo"></div>
                </div>
                <div class="modal-footer">
                    <button type="button" class="btn btn-sm btn-secondary" data-bs-dismiss="modal">Fechar</button>
                </div>
            </div>
        </div>
    </div>
</body>
</html>