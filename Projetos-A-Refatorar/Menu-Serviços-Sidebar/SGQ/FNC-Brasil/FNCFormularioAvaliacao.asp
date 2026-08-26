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

    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.0.2/dist/css/bootstrap.min.css" rel="stylesheet" crossorigin="anonymous" />
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

        #btnVoltarPagina:disabled,
        #btnProximaPagina:disabled {
            cursor: not-allowed;
            pointer-events: auto;
        }

        th,
        td {
            white-space: nowrap;
            vertical-align: top;
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

        .readonly-alert {
            background: #f8f9fa;
            border: 1px dashed #ced4da;
            border-radius: 8px;
            color: #6c757d;
            padding: 12px 14px;
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

        .table thead {
            vertical-align: middle !important;
            position: sticky;
            inset-block-start: 0;
            background: #f8f9fa;
            z-index: 1;
        }

        .table td,
        .table th {
            padding: 8px 12px !important;
        }

        .table th {
            color: #495057;
            font-weight: 600;
        }

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

        .just-validate-error-label {
            color: #f87171 !important;
            display: block;
            font-size: .8rem;
            margin-top: 3px;
        }

        .just-validate-error-field {
            border-color: #f87171 !important;
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
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.0.2/dist/js/bootstrap.bundle.min.js" crossorigin="anonymous"></script>
    <script src="https://cdn.jsdelivr.net/npm/select2@4.1.0-rc.0/dist/js/select2.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/jquery.mask/1.14.15/jquery.mask.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/just-validate@latest/dist/just-validate.production.min.js"></script>

    <script src="/corporativo/components/toast.js?v=<%=Now()%>"></script>
    <script src="/corporativo/components/confirmation.js?v=<%=Now()%>"></script>
    <script src="/corporativo/components/upload/fileUpload.js?v=<%=Now()%>"></script>
    <script src="../Diagramas/Ishikawa/ishikawa.js?v=<%=Now()%>"></script>
    <script src="../Diagramas/5W2H/5W2H.js?v=<%=Now()%>"></script>
    <script src="/corporativo/funcoes/funcoesJAVASCRIPT.js?v=<%=Now()%>"></script>

    <script>
        let id = 0;
        let acao = '';
        let abaAtual = 1;
        let totalAbas = 9;
        let limitesNQA = {};

        function repoUrl() {
            return 'FNCRepositorio.asp?v=<%=Now()%>';
        }

        function inicializarSelect2() {
            $('.select2').select2({
                theme: 'bootstrap-5',
                width: '100%',
                dropdownCssClass: 'z-index-9999'
            });
        }

        function navegarParaGrade() {
            window.location.href = 'FNCGrade.asp';
        }

        function exibirLoading(exibir) {
            if (exibir) {
                $('#loadingOverlay').removeClass('d-none').addClass('d-flex');
            } else {
                $('#loadingOverlay').removeClass('d-flex').addClass('d-none');
            }
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
                $button.find('.tab-icon').attr('class', 'tab-icon bi bi-check-circle-fill');
            }
        }

        function marcarAbaComoInvalida(numeroAba) {
            const $button = $(`.tab-button[data-tab="${numeroAba}"]`);

            if ($button.length) {
                $button.addClass('invalid').removeClass('completed');
                $button.find('.tab-icon').attr('class', 'tab-icon bi bi-exclamation-circle-fill');
            }
        }
           
        function irProximo(){
            navegarParaAba(abaAtual + 1);
        }

        function irAnterior(){
            navegarParaAba(abaAtual - 1);
        }

        function anexarEventos() {
            $('#formularioAvaliacaoFNC').on('submit', function(e) {
                e.preventDefault();
                return false;
            });

            $('.tab-button').on('click', function() {
                navegarParaAba(parseInt($(this).data('tab')));
            });

            $('#btnAprovarAvaliacao').on('click', function() {
                avaliarFNC(true);
            });

            $('#btnRejeitarAvaliacao').on('click', function() {
                avaliarFNC(false);
            });

            $('#btnAprovarResposta').on('click', function() {
                avaliarFNC(true);
            });

            $('#btnRejeitarResposta').on('click', function() {
                avaliarFNC(false);
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

                    limitesNQA = response.dados || {};
                },
                error: function(xhr, status, error) {
                    console.error('Erro ao carregar limites NQA:', error);
                }
            });
        }

        function carregarDadosFNC() {
            exibirLoading(true);

            $.ajax({
                url: repoUrl(),
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

                    preencherDados(response.dados);
                },
                error: function(xhr, status, error) {
                    console.error('Erro ao carregar dados:', error);
                    Toast.error('Erro ao carregar dados da FNC');
                },
                complete: function() {
                    exibirLoading(false);
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
                        $select.append($('<option>', {
                            value: item.id,
                            text: item.descricao
                        }));
                    });
                },
                error: function(xhr, status, error) {
                    console.error('Erro ao carregar defeitos:', error);
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
                    if (!response.sucesso) {
                        console.error('Erro ao carregar disposições: ' + response.mensagem);
                        return;
                    }

                    response.dados.forEach(function(item) {
                        $select.append($('<option>', {
                            value: item.id,
                            text: item.descricao
                        }));
                    });
                },
                error: function(xhr, status, error) {
                    console.error('Erro ao carregar disposições:', error);
                }
            });
        }

        function preencherDados(dados) {
            renderizarNotas(dados.notas || []);
            renderizarAnexos(dados.anexos || []);
            renderizarObservacoes(dados.observacoes || []);
            renderizarRespostas(dados.respostas || []);
            renderizarAprovacoes(dados.aprovacoes || []);

            $.when(
                carregarDisposicoes(),
                carregarDefeitos(dados.tipoDefeito)
            ).then(function() {
                $('#codPlanta').val(dados.codPlanta);
                $('#nomeCliente').val(dados.nomeCliente);
                $('#codCliente').val(dados.codCliente);
                $('#codLoja').val(dados.codLoja);
                $('#cnpj').val(formatarCNPJ(dados.cnpj || ''));
                atualizarTotaisReclamados(dados.notas || []);
                $('#descricaoDisposicao').val(dados.descricaoDisposicao);
                $('#descricaoDefeito').val(dados.descricaoDefeito);
                $('#dataInclusao').val(dados.dataInclusao);
                $('#dataEnvioCliente').val(dados.dataEnvioCliente);
                $('#eficacia').val(dados.eficacia);
                $('#vendedor').val(`${dados.codVendedor} - ${dados.nomeVendedor}`)

                setTimeout(function() {
                    $('#tipoDefeito').val(dados.tipoDefeito).trigger('change.select2');
                    $('#disposicaoId').val(dados.disposicaoId).trigger('change.select2');
                    $('#defeitoId').val(dados.defeitoId).trigger('change.select2');
                }, 300);

                if (dados.tipoDefeito == 'PRO') {
                    exibirCoeficienteNQA(parseFloat(dados.coeficienteNQA || 0));
                }

                renderizarIshikawa(dados.causasIshikawa || [], dados.defeitoIshikawa || dados.descricaoTipoDefeito);
                renderizar5W2H(dados.acoes5W2H || [], dados.assunto5W2H || dados.descricaoTipoDefeito);
            });
        }

        function validarAcao(flagAprovar) {
            const observacao = ($('#observacao').val() || '').trim();
            const resposta = ($('#resposta').val() || '').trim();
            const justificativa = ($('#justificativa').val() || '').trim();

            $('.just-validate-error-label').remove();
            $('.just-validate-error-field').removeClass('just-validate-error-field');

            if (acao === 'AVALIAR' && !flagAprovar && !observacao) {
                $('#observacao')
                    .addClass('just-validate-error-field')
                    .after('<div class="just-validate-error-label">Observação obrigatória para rejeição</div>');

                marcarAbaComoInvalida(1);
                navegarParaAba(1);
                Toast.warning('Observação obrigatória para rejeição');
                return false;
            }

            if (acao === 'RESPONDER') {
                if (!resposta) {
                    $('#resposta')
                        .addClass('just-validate-error-field')
                        .after('<div class="just-validate-error-label">Resposta é obrigatória</div>');

                    marcarAbaComoInvalida(1);
                    navegarParaAba(1);
                    Toast.warning('Resposta é obrigatória');
                    return false;
                }

                if (flagAprovar && !justificativa) {
                    $('#justificativa')
                        .addClass('just-validate-error-field')
                        .after('<div class="just-validate-error-label">Plano de ação é obrigatório quando procede</div>');

                    marcarAbaComoInvalida(1);
                    navegarParaAba(1);
                    Toast.warning('Plano de ação é obrigatório quando procede');
                    return false;
                }
            }

            marcarAbaComoCompleta(1);
            return true;
        }

        function avaliarFNC(flagAprovar) {
            if (!validarAcao(flagAprovar)) {
                return;
            }

            const observacao = $('#observacao').val() || '';
            const resposta = $('#resposta').val() || '';
            const justificativa = $('#justificativa').val() || '';
            const anexos = typeof uploadAnexos !== 'undefined' && uploadAnexos ? uploadAnexos.getFiles() : [];

            const causasIshikawa = ishikawa
                ? ishikawa.getData().map(causa => ({
                    ...causa,
                    causaRaiz: causa.causaRaiz ? 1 : 0
                }))
                : [];

            const acoes5W2H = W2H ? W2H.getData() : [];
            const defeitoIshikawa = ishikawa ? ishikawa.opts.problema : '';

            exibirLoading(true);

            $.ajax({
                url: repoUrl(),
                type: 'POST',
                dataType: 'json',
                data: {
                    id: id,
                    acao: acao === 'AVALIAR' ? 'avaliar' : 'responder',
                    flagAprovar: flagAprovar ? 1 : 0,
                    observacao: observacao.replaceAll("'", '"'),
                    resposta: resposta.replaceAll("'", '"'),
                    justificativa: justificativa.replaceAll("'", '"'),
                    defeitoIshikawa: defeitoIshikawa,
                    anexos: JSON.stringify(anexos),
                    causasIshikawa: JSON.stringify(causasIshikawa),
                    assunto5W2H: W2H ? W2H.opts.assunto : '',
                    acoes5W2H: JSON.stringify(acoes5W2H)
                },
                success: function(response) {
                    if (!response.sucesso) {
                        Toast.error(response.mensagem);
                        return;
                    }

                    Toast.success(acao === 'AVALIAR'
                        ? 'Avaliação registrada com sucesso!'
                        : 'Resposta registrada com sucesso!'
                    );

                    setTimeout(function() {
                        navegarParaGrade();
                    }, 1000);
                },
                error: function(xhr, status, error) {
                    console.error('Erro:', error);
                    Toast.error('Erro ao processar solicitação');
                },
                complete: function() {
                    exibirLoading(false);
                }
            });
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

        function calcularTotalQuantidadeReclamada(listaNotas) {
            return (listaNotas || []).reduce((total, nota) => {
                return total + obterQuantidadeReclamadaNumerica(nota.quantidadeReclamada);
            }, 0);
        }

        function calcularTotalValorReclamado(listaNotas) {
            return (listaNotas || []).reduce((total, nota) => {
                return total + calcularValorReclamadoNota(nota);
            }, 0);
        }

        function atualizarTotaisReclamados(listaNotas) {
            const quantidadeFormatada = formatarQuantidadeReclamada(calcularTotalQuantidadeReclamada(listaNotas)) || '0';
            const valorFormatado = formatarValorReclamado(calcularTotalValorReclamado(listaNotas));
            $('#quantidadeReclamada').val(quantidadeFormatada);
            $('#valorReclamado').val(valorFormatado);
        }

        function renderizarNotas(listaNotas) {
            listaNotas = listaNotas || [];
            $('#contadorNotas, #menuContadorNotas').text(listaNotas.length);
            atualizarTotaisReclamados(listaNotas);
            
            if (listaNotas.length === 0) {
                $('#notasListaContainer').html(`
                    <div class="text-center py-5">
                        <p class="text-muted mt-3 mb-0">Nenhuma nota vinculada.</p>
                    </div>
                `);
                return;
            }
            let html = `
                <div style="overflow: auto; max-height: 45vh;" class="table-responsive">
                    <table class="table table-sm table-hover mb-0">
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
                            </tr>
                        </thead>
                        <tbody>
            `;
            listaNotas.forEach(nota => {
                nota.valorReclamado = formatarValorReclamado(calcularValorReclamadoNota(nota));
                html += `
                    <tr>
                        <td>${nota.numeroPedido || '-'}</td>
                        <td>${nota.itemPedido || '-'}</td>
                        <td>${nota.referencia || '-'}</td>
                        <td>${nota.quantidade || '-'}</td>
                        <td>${nota.quantidadeReclamada || '-'}</td>
                        <td class="text-end">${nota.precoUnitario || '-'}</td>
                        <td class="text-end">${nota.valorReclamado}</td>
                        <td>${nota.dataEmissaoNota || '-'}</td>
                        <td>${nota.notaFiscal || '-'}</td>
                        <td>${nota.serieNotaFiscal || '-'}</td>
                    </tr>
                `;
            });
            html += `
                        </tbody>
                    </table>
                </div>
            `;
            $('#notasListaContainer').html(html);
        }
        
        function renderizarAnexos(listaAnexos) {
            $('#menuContadorAnexos').text(listaAnexos.length);

            const $container = $('#anexosListaContainer');

            if (listaAnexos.length === 0) {
                $container.html(``);
                return;
            }

            let html = '<div class="anexos-list mt-4">';

            listaAnexos.forEach(anexo => {
                html += `
                    <div class="anexo-item d-flex align-items-center">
                        <i class="bi bi-file-earmark anexo-icon"></i>
                        <div class="anexo-info">
                            <div class="anexo-nome">${anexo.nome || '-'}</div>
                            <div class="anexo-data">Adicionado por <strong>${anexo.usuario}</strong> em: ${anexo.dataCriacao}</div>
                        </div>
                        <a href="/Clientes_Docs/${anexo.caminho}" target="_blank" class="btn btn-sm btn-primary">
                            <i class="bi bi-box-arrow-up-right me-1"></i> Abrir
                        </a>
                    </div>
                `;
            });

            html += '</div>';
            $container.html(html);
        }

        function renderizarObservacoes(listaObservacoes) {
            $('#contadorObservacoes, #menuContadorObservacoes').text(listaObservacoes.length);

            const $container = $('#observacoesListaContainer');

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
                                <th>#</th>
                                <th>Descrição</th>
                                <th>Incluída por</th>
                                <th>Data inclusão</th>
                            </tr>
                        </thead>
                        <tbody>
            `;

            listaObservacoes.forEach(observacao => {
                html += `
                    <tr>
                        <td>${observacao.id || '-'}</td>
                        ${criarCelulaDetalhe(observacao.descricao, 'Observação')}
                        <td>${observacao.loginUsuarioInclusao || '-'}</td>
                        <td>${observacao.dataCriacao || '-'}</td>
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

        function renderizarRespostas(listaRespostas) {
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

            if (listaRespostas.length === 0) {
                html += `
                    <tr>
                        <td colspan="7" class="text-center text-muted py-4">Nenhuma resposta registrada.</td>
                    </tr>
                `;
            }

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

        function renderizarAprovacoes(listaAprovacoes) {
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

            if (listaAprovacoes.length === 0) {
                html += `
                    <tr>
                        <td colspan="5" class="text-center text-muted py-4">Nenhuma aprovação registrada.</td>
                    </tr>
                `;
            }

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

        function renderizarIshikawa(causasIshikawa, descricaoProblema) {
            ishikawa = new Ishikawa({
                el: '#ishikawa-container',
                problema: descricaoProblema,
                readOnly: acao === 'AVALIAR'
            });

            ishikawa.setData((causasIshikawa || []).map(causa => ({
                ...causa,
                causaRaiz: causa.causaRaiz == '1'
            })));
        }

        function renderizar5W2H(acoes5W2H, descricaoProblema) {
            W2H = new W2H({
                el: '#5W2H-container',
                assunto: descricaoProblema,
                readOnly: acao === 'AVALIAR'
            });

            W2H.setData(acoes5W2H || []);
        }

        function exibirCoeficienteNQA(coeficienteNQA) {
            const box = document.getElementById('nqa-inline');
            
            if (!box || !coeficienteNQA) {
                return;
            }

            let state, badge, detail;
            if(coeficienteNQA < limitesNQA.critico) {
                state = 's-ok';
                badge = 'Aprovado';
                detail = `Menor do que ${limitesNQA.critico}%`;
            } else if ( (coeficienteNQA >= limitesNQA.critico) && (coeficienteNQA < parseFloat(limitesNQA.grave)) ) {
                state = 's-fail';
                badge = 'Crítico';
                detail = `Maior do que ${limitesNQA.critico}%`;
            } else {
                state = 's-fail';
                badge = 'Grave';
                detail = `Maior do que ${limitesNQA.grave}%`;
            }

            $('#nq-val').text(formatarMoeda(coeficienteNQA) + '%');
            $('#nq-badge').text(badge);
            $('#nq-detail').text(detail);

            box.classList.remove('s-ok', 's-warn', 's-fail');
            box.classList.add('visible', state);
        }

        function configurarTelaPorAcao() {
            if (acao === 'AVALIAR') {
                $('#titulo').text(`Avaliar FNC #${id}`);
                $('#tituloAcao').text('Avaliação da resposta');
                $('#subtituloAcao').text('Registre a aprovação ou rejeição da resposta enviada pelo responsável.');
                $('#containerAvaliar').removeClass('d-none');
                $('#containerResponder').addClass('d-none');
                $('#labelAbaAcao').text('Avaliação');
                return true;
            }

            if (acao === 'RESPONDER') {
                $('#titulo').text(`Responder FNC #${id}`);
                $('#tituloAcao').text('Resposta da FNC');
                $('#subtituloAcao').text('Informe a resposta da análise e, quando procedente, o plano de ação.');
                $('#containerResponder').removeClass('d-none');
                $('#containerAvaliar').addClass('d-none');
                $('#labelAbaAcao').text('Resposta');
                return true;
            }

            Toast.error('Ação inválida');
            return false;
        }

        $(document).ready(function() {
            $('.tab-button').prop('disabled', true);

            const urlParams = new URLSearchParams(window.location.search);
            const paramId = urlParams.get('id');
            const paramAcao = urlParams.get('acao');

            if (!paramId) {
                Toast.error('ID inválido');
                return;
            }

            if (!paramAcao) {
                Toast.error('Ação inválida');
                return;
            }

            id = paramId;
            acao = paramAcao.toUpperCase();

            if (!configurarTelaPorAcao()) {
                return;
            }

            inicializarSelect2();
            anexarEventos();
            navegarParaAba(1);

            $.when(carregarLimitesNQA()).always(function() {
                carregarDadosFNC();
            });

            $('.tab-button').prop('disabled', false);
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
                        <div class="tab-group-title">Ação requerida</div>

                        <button type="button" class="tab-button active" data-tab="1">
                            <i class="tab-icon bi bi-check2-square"></i>
                            <span class="tab-title" id="labelAbaAcao">Ação</span>
                        </button>

                        <div class="tab-group-title">Consulta</div>

                        <button type="button" class="tab-button" data-tab="2">
                            <i class="tab-icon bi bi-clipboard-check"></i>
                            <span class="tab-title">Cadastro</span>
                        </button>

                        <button type="button" class="tab-button" data-tab="3">
                            <i class="tab-icon bi bi-receipt"></i>
                            <span class="tab-title">Notas</span>
                            <span class="tab-meta" id="menuContadorNotas">0</span>
                        </button>

                        <div class="tab-group-title">Ferramentas de Qualidade</div>

                        <button type="button" class="tab-button" data-tab="4">
                            <i class="tab-icon bi bi-diagram-3"></i>
                            <span class="tab-title">Ishikawa</span>
                        </button>

                        <button type="button" class="tab-button" data-tab="5">
                            <i class="tab-icon bi bi-list-check"></i>
                            <span class="tab-title">Plano 5W2H</span>
                        </button>

                        <div class="tab-group-title">Adicionais</div>

                        <button type="button" class="tab-button" data-tab="6">
                            <i class="tab-icon bi bi-paperclip"></i>
                            <span class="tab-title">Anexos</span>
                            <span class="tab-meta" id="menuContadorAnexos">0</span>
                        </button>

                        <button type="button" class="tab-button" data-tab="7">
                            <i class="tab-icon bi bi-chat-dots"></i>
                            <span class="tab-title">Observações</span>
                            <span class="tab-meta" id="menuContadorObservacoes">0</span>
                        </button>

                        <div class="tab-group-title">Histórico</div>

                        <button type="button" class="tab-button" data-tab="8">
                            <i class="tab-icon bi bi-reply"></i>
                            <span class="tab-title">Respostas</span>
                            <span class="tab-meta" id="menuContadorRespostas">0</span>
                        </button>

                        <button type="button" class="tab-button" data-tab="9">
                            <i class="tab-icon bi bi-shield-check"></i>
                            <span class="tab-title">Aprovações</span>
                            <span class="tab-meta" id="menuContadorAprovacoes">0</span>
                        </button>
                    </div>
                </div>

                <div class="tab-content-wrapper">
                    <form id="formularioAvaliacaoFNC" name="formularioAvaliacaoFNC" novalidate>
                        <div class="tab-content active" data-tab-content="1">
                            <div class="section-header">
                                <div class="section-title" id="tituloAcao">Ação da FNC</div>
                                <p class="section-subtitle" id="subtituloAcao"></p>
                            </div>

                            <div id="containerAvaliar" class="d-none">
                                <div class="row g-3">
                                    <div class="col-12">
                                        <label for="observacao" class="form-label">
                                            Observações
                                            <span class="text-muted fw-normal">(obrigatória para rejeição)</span>
                                        </label>
                                        <textarea class="form-control form-control-sm" id="observacao" rows="5"></textarea>
                                    </div>

                                    <div class="col-12 d-flex gap-2 justify-content-end mt-3">
                                        <button type="button" class="btn btn-sm btn-danger" id="btnRejeitarAvaliacao">
                                            <i class="bi bi-x-lg me-1"></i>Rejeitar
                                        </button>

                                        <button type="button" class="btn btn-sm btn-success" id="btnAprovarAvaliacao">
                                            <i class="bi bi-check-lg me-1"></i>Aprovar
                                        </button>
                                    </div>
                                </div>
                            </div>

                            <div id="containerResponder" class="d-none">
                                <div class="row g-3">
                                    <div class="col-12">
                                        <label for="resposta" class="form-label">
                                            Resposta <span class="required-field">*</span>
                                        </label>
                                        <textarea class="form-control form-control-sm" id="resposta" rows="5"></textarea>
                                    </div>

                                    <div class="col-12">
                                        <label for="justificativa" class="form-label">
                                            Plano de ação
                                            <span class="text-muted fw-normal">(obrigatório quando procede)</span>
                                        </label>
                                        <textarea class="form-control form-control-sm" id="justificativa" rows="5"></textarea>
                                    </div>

                                    <div class="col-12 d-flex gap-2 justify-content-end mt-3">
                                        <button type="button" class="btn btn-sm btn-danger" id="btnRejeitarResposta">
                                            <i class="bi bi-x-lg me-1"></i>Não procede
                                        </button>

                                        <button type="button" class="btn btn-sm btn-success" id="btnAprovarResposta">
                                            <i class="bi bi-check-lg me-1"></i>Procede
                                        </button>
                                    </div>
                                </div>
                            </div>

                            <div class="readonly-alert mt-3">
                                <i class="bi bi-info-circle me-1"></i>
                                Consulte as demais abas para revisar cadastro, notas, anexos, observações e históricos antes de concluir.
                            </div>
                        </div>

                        <div class="tab-content" data-tab-content="2">
                            <div class="section-header">
                                <div class="section-title">Cadastro da reclamação</div>
                                <p class="section-subtitle">Dados principais da FNC em modo consulta.</p>
                            </div>

                            <div class="nqa-inline" id="nqa-inline">
                                <span class="nq-row-label">Quociente NQA:</span>
                                <span class="nq-val" id="nq-val">—</span>
                                <span class="nq-sep"></span>
                                <span class="nq-badge" id="nq-badge">—</span>
                                <span class="nq-detail" id="nq-detail"></span>
                            </div>

                            <div class="row g-3">
                                <div class="col-12 col-md-6">
                                    <label for="nomeCliente" class="form-label">Nome do Cliente</label>
                                    <input type="text" class="form-control form-control-sm" id="nomeCliente" disabled>
                                </div>

                                <div class="col-12 col-md-2">
                                    <label for="codCliente" class="form-label">Cód. Cliente</label>
                                    <input type="text" class="form-control form-control-sm" id="codCliente" disabled>
                                </div>

                                <div class="col-12 col-md-2">
                                    <label for="codLoja" class="form-label">Loja</label>
                                    <input type="text" class="form-control form-control-sm" id="codLoja" disabled>
                                </div>

                                <div class="col-12 col-md-2">
                                    <label for="cnpj" class="form-label">CNPJ</label>
                                    <input type="text" class="form-control form-control-sm" id="cnpj" disabled>
                                </div>

                                <div class="col-12 col-md-4">
                                    <label for="dataInclusao" class="form-label">Data de Inclusão</label>
                                    <input disabled type="date" class="form-control form-control-sm" id="dataInclusao">
                                </div>

                                <div class="col-12 col-md-4">
                                    <label for="dataEnvioCliente" class="form-label">Data de Envio ao Cliente</label>
                                    <input disabled type="date" class="form-control form-control-sm" id="dataEnvioCliente">
                                </div>

                                <div class="col-12 col-md-4">
                                    <label for="tipoDefeito" class="form-label">Tipo de Defeito</label>
                                    <select disabled class="form-select form-select-sm select2" id="tipoDefeito">
                                        <option value="">Selecione o tipo de defeito</option>
                                        <option value="ADM">Administrativo</option>
                                        <option value="PRO">Industrial</option>
                                    </select>
                                </div>

                                <div class="col-12 col-md-4">
                                    <label for="defeitoId" class="form-label">Classificação Defeito</label>
                                    <select disabled class="form-select form-select-sm select2" id="defeitoId">
                                        <option value="">Selecione um defeito</option>
                                    </select>
                                </div>

                                <div class="col-12 col-md-4">
                                    <label for="disposicaoId" class="form-label">Código Disposição</label>
                                    <select disabled class="form-select form-select-sm select2" id="disposicaoId">
                                        <option value="">Selecione a disposição</option>
                                    </select>
                                </div>
                                
                                <div class="col-12 col-md-4">
                                    <label for="vendedor" class="form-label">Vendedor <span class="required-field">*</span></label>
                                    <input type="text" class="form-control form-control-sm" id="vendedor" name="vendedor" disabled>
                                </div>

                                <div class="col-12">
                                    <label for="descricaoDisposicao" class="form-label">Disposição</label>
                                    <textarea disabled class="form-control form-control-sm" id="descricaoDisposicao" rows="4"></textarea>
                                </div>

                                <div class="col-12">
                                    <label for="descricaoDefeito" class="form-label">Descrição Defeito</label>
                                    <textarea disabled class="form-control form-control-sm" id="descricaoDefeito" rows="4"></textarea>
                                </div>

                                <div class="col-12">
                                    <label for="eficacia" class="form-label">Eficácia/Observação</label>
                                    <textarea disabled class="form-control form-control-sm" id="eficacia" rows="4"></textarea>
                                </div>
                            </div>
                        </div>

                        <div class="tab-content" data-tab-content="3">
                            <div class="section-header">
                                <div class="section-title">
                                    Notas vinculadas <span class="badge bg-secondary ms-1" id="contadorNotas">0</span>
                                </div>
                                <p class="section-subtitle">Notas fiscais relacionadas à reclamação.</p>
                            </div>

                            <div class="row g-3 mb-3">
                                <div class="col-lg-4 col-md-6">
                                    <label for="quantidadeReclamada" class="form-label">Quantidade Reclamada Total</label>
                                    <input disabled type="text" class="form-control form-control-sm" id="quantidadeReclamada" value="0">
                                </div>
                                <div class="col-lg-4 col-md-6">
                                    <label for="valorReclamado" class="form-label">Valor Reclamado Total</label>
                                    <input disabled type="text" class="form-control form-control-sm" id="valorReclamado" value="0,00">
                                </div>
                            </div>
                            <div id="notasListaContainer"></div>
                        </div>

                        <div class="tab-content" data-tab-content="4">
                            <div class="section-header">
                                <div class="section-title">Diagrama de Ishikawa</div>
                                <p class="section-subtitle">Análise de causa raiz do problema.</p>
                            </div>

                            <div id="ishikawa-container"></div>
                        </div>

                        <div class="tab-content" data-tab-content="5">
                            <div class="section-header">
                                <div class="section-title">Plano de ação 5W2H</div>
                                <p class="section-subtitle">Ações, responsáveis, prazos e tratativas da não conformidade.</p>
                            </div>

                            <div id="5W2H-container"></div>
                        </div>

                        <div class="tab-content" data-tab-content="6">
                            <div class="section-header">
                                <div class="section-title">Anexos</div>
                                <p class="section-subtitle">Evidências, documentos e arquivos relacionados à FNC.</p>
                            </div>

                            <div class="row g-3">
                                <file-upload id="uploadAnexos" file-code="FNC" max-files="10" multiple></file-upload>
                                <div id="anexosListaContainer"></div>
                            </div>
                        </div>

                        <div class="tab-content" data-tab-content="7">
                            <div class="section-header">
                                <div class="section-title">
                                    Observações <span class="badge bg-secondary ms-1" id="contadorObservacoes">0</span>
                                </div>
                                <p class="section-subtitle">Comentários e registros adicionais da FNC.</p>
                            </div>

                            <div id="observacoesListaContainer"></div>
                        </div>

                        <div class="tab-content" data-tab-content="8">
                            <div class="section-header">
                                <div class="section-title">
                                    Histórico de respostas <span class="badge bg-secondary ms-1" id="contadorRespostas">0</span>
                                </div>
                                <p class="section-subtitle">Consulta dos retornos registrados para esta FNC.</p>
                            </div>

                            <div id="respostasListaContainer"></div>
                        </div>

                        <div class="tab-content" data-tab-content="9">
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
