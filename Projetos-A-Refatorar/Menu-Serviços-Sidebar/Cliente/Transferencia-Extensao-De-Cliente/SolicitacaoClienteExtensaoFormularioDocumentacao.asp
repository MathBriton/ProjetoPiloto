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
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />

    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.1/font/bootstrap-icons.css" />
    <link
        href="https://cdn.jsdelivr.net/npm/bootstrap@5.0.2/dist/css/bootstrap.min.css"
        rel="stylesheet"
        integrity="sha384-EVSTQN3/azprG1Anm3QDgpJLIm9Nao0Yz1ztcQTwFspd3yD65VohhpuuCOmLASjC"
        crossorigin="anonymous"
    />
    <link href="https://cdn.jsdelivr.net/npm/select2@4.1.0-rc.0/dist/css/select2.min.css" rel="stylesheet" />
    <link href="https://cdn.jsdelivr.net/npm/select2-bootstrap-5-theme@1.3.0/dist/select2-bootstrap-5-theme.min.css" rel="stylesheet" />

    <style>
        :root {
            --surface: #ffffff;
            --surface-muted: #f8f9fa;
            --border: #e9ecef;
            --border-strong: #dee2e6;
            --text: #1f2937;
            --text-muted: #6c757d;
            --radius: 10px;
        }

        body {
            background-color: #f5f5f5;
            font-size: .9rem;
        }

        .container {
            max-width: 85%;
        }

        .btn-back {
            font-size: 1.25rem;
        }

        .loading-overlay {
            align-items: center;
            background: rgba(255, 255, 255, .8);
            height: 100%;
            justify-content: center;
            left: 0;
            position: fixed;
            top: 0;
            width: 100%;
            z-index: 9999;
        }

        .form-container {
            background: var(--surface);
            border-radius: var(--radius);
            box-shadow: 0 2px 4px rgba(0, 0, 0, .1);
            overflow: hidden;
        }

        .form-layout {
            display: grid;
            grid-template-columns: 270px minmax(0, 1fr);
        }

        .form-label {
            font-weight: 600;
        }

        .form-content {
            overflow: hidden;
        }

        .required-field {
            color: #dc3545;
        }

        .tabs-container {
            background-color: var(--surface-muted);
            border-right: 1px solid var(--border);
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
            background: var(--border);
            border-radius: 999px;
            color: var(--text-muted);
            font-size: 10px;
            font-weight: 700;
            margin-left: auto;
            padding: 1px 7px;
        }

        .tab-button.active .tab-meta {
            background: #d7eaff;
            color: #0d6efd;
        }

        .tab-meta.meta-pendente { background: #fff3cd; color: #997404; }
        .tab-meta.meta-invalido { background: #f8d7da; color: #dc3545; }
        .tab-meta.meta-ok       { background: #d1e7dd; color: #198754; }

        .tab-content-wrapper {
            background: var(--surface);
            display: flex;
            flex-direction: column;
            min-width: 0;
        }

        .tab-content {
            animation: fadeIn .25s ease;
            display: none;
            flex: 1;
            height: 72vh;
            overflow: auto;
            padding: 28px 30px;
        }

        .tab-content.active {
            display: block;
        }

        .section-header {
            border-bottom: 1px solid var(--border);
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
            color: var(--text-muted);
            font-size: .825rem;
            margin: 0;
        }

        .readonly-alert {
            background: var(--surface-muted);
            border: 1px dashed #ced4da;
            border-radius: 8px;
            color: var(--text-muted);
            padding: 12px 14px;
        }

        /* ----- Barra de acao principal (rodape fixo) ----- */
        .action-bar {
            align-items: center;
            background: var(--surface);
            border-top: 1px solid var(--border-strong);
            bottom: 0;
            display: flex;
            gap: 10px;
            justify-content: end;
            padding: 12px 20px;
            position: sticky;
            z-index: 10;
        }

        .action-bar .action-buttons {
            display: flex;
            gap: 8px;
            flex-shrink: 0;
        }

        /* ----- Resumo de progresso ----- */
        .aval-resumo {
            display: flex;
            flex-wrap: wrap;
            gap: 8px;
            margin-bottom: 18px;
        }

        .resumo-chip {
            align-items: center;
            background: #f1f3f5;
            border-radius: 999px;
            color: #495057;
            display: inline-flex;
            font-size: .78rem;
            font-weight: 600;
            gap: 6px;
            padding: .32rem .75rem;
        }

        .resumo-chip i { font-size: .95rem; }
        .resumo-chip.chip-valido   { background: #d1e7dd; color: #198754; }
        .resumo-chip.chip-invalido { background: #f8d7da; color: #dc3545; }
        .resumo-chip.chip-pendente { background: #fff3cd; color: #997404; }

        /* ----- Grade de anexos (duas colunas) ----- */
        /* ===== Design alternativo: lista horizontal de documentos ===== */
        .anexos-list {
            display: flex;
            flex-direction: column;
            gap: 10px;
        }

        .anexo-item {
            background: var(--surface);
            border: 1px solid #e8ebef;
            border-radius: 12px;
            box-shadow: inset 3px 0 0 #d7dbe0;     
            display: flex;
            flex-direction: column;
            padding: 12px 16px 12px 18px;
            transition: background .15s, border-color .15s, box-shadow .15s;
        }

        .anexo-item:hover { border-color: #cfd6df; }

        .anexo-item.is-valido   { box-shadow: inset 3px 0 0 #198754; }  
        .anexo-item.is-invalido { box-shadow: inset 3px 0 #dc3545; }  

        .anexo-main {
            align-items: center;
            display: flex;
            gap: 14px;
        }

        /* Abrir o documento: icone + nome formam o link */
        .anexo-link {
            align-items: center;
            color: inherit;
            display: flex;
            flex: 1;
            gap: 12px;
            min-width: 0;
            text-decoration: none;
        }

        .anexo-link:hover .anexo-nome { text-decoration: underline; }

        .anexo-icon-wrap {
            align-items: center;
            border-radius: 9px;
            display: flex;
            flex-shrink: 0;
            font-size: 1.2rem;
            height: 38px;
            justify-content: center;
            width: 38px;
        }

        .ico-pdf,
        .ico-img,
        .ico-doc,
        .ico-xls,
        .ico-generico { background: #f1f3f5; color: #6b7280; }

        .anexo-nome {
            color: var(--text);
            flex: 1;
            font-weight: 600;
            line-height: 1.3;
            min-width: 0;
            overflow: hidden;
            text-overflow: ellipsis;
            white-space: nowrap;
        }

        .anexo-ext-ico {
            color: var(--text-muted);
            flex-shrink: 0;
            font-size: .85rem;
            opacity: .35;
            transition: opacity .15s, color .15s;
        }

        .anexo-link:hover .anexo-ext-ico { color: #0d6efd; opacity: 1; }

        .anexo-toggle {
            flex-shrink: 0;
        }

        .anexo-motivo {
            margin-top: 12px;
            overflow: hidden;
        }

        .anexo-motivo:not(.d-none) {
            animation: motivoIn .22s ease;
        }

        @keyframes motivoIn {
            from { max-height: 0; opacity: 0; transform: translateY(-4px); }
            to   { max-height: 240px; opacity: 1; transform: translateY(0); }
        }

        @media (max-width: 575px) {
            .anexo-main { flex-wrap: wrap; }
            .anexo-link { flex: 1 1 100%; }
            .anexo-toggle { flex: 1 1 100%; width: 100%; }
            .anexo-toggle .btn { flex: 1; }
        }

        .bloco-parecer {
            border-top: 1px solid var(--border);
            margin-top: 24px;
            padding-top: 20px;
        }

        .just-validate-error-label,
        .erro-validacao {
            color: #f87171 !important;
            display: block;
            font-size: .8rem;
            margin-top: 3px;
        }

        .just-validate-error-field {
            border-color: #f87171 !important;
        }

        th, td {
            vertical-align: top;
            white-space: nowrap;
        }

        @keyframes fadeIn {
            from { opacity: 0; transform: translateY(6px); }
            to   { opacity: 1; transform: translateY(0); }
        }

        @media (max-width: 1400px) {
            .container { max-width: 95%; }
            .form-layout { grid-template-columns: 1fr; }

            .tabs-container {
                border-bottom: 1px solid var(--border);
                border-right: none;
                overflow-x: auto;
            }

            .tabs {
                flex-direction: row;
                min-width: max-content;
                padding: 8px;
            }

            .tab-group-title { display: none; }

            .tab-button {
                justify-content: center;
                min-width: 150px;
                text-align: center;
                width: auto;
            }

            .tab-title { flex: initial; }
        }

        @media (max-width: 575px) {
            .action-bar { flex-direction: column; align-items: stretch; }
            .action-bar .action-buttons .btn { flex: 1; }
        }
    </style>

    <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.0.2/dist/js/bootstrap.bundle.min.js" integrity="sha384-MrcW6ZMFYlzcLA8Nl+NtUVF0sA7MsXsP1UyJoMp4YLEuNSfAP+JcXn/tWtIaxVXM"crossorigin="anonymous"></script>
    <script src="https://cdn.jsdelivr.net/npm/select2@4.1.0-rc.0/dist/js/select2.min.js"></script>
    <script src="/corporativo/components/toast.js?v=<%=Now()%>"></script>
    <script src="/corporativo/components/confirmation.js?v=<%=Now()%>"></script>
    <script src="/corporativo/funcoes/funcoesJAVASCRIPT.js?v=<%=Now()%>"></script>

    <script>
        const STATUS_ANEXO = {
            PENDENTE: 'PENDENTE',
            VALIDO: 'VALIDO',
            INVALIDO: 'INVALIDO'
        };

        let id = 0;
        let acao = '';
        let abaAtual = 1;
        let totalAbas = 2;
        let dadosSolicitacao = null;
        let listaAnexos = [];
        let avaliacoesAnexos = {};

        function repoUrl() {
            return 'SolicitacaoClienteExtensaoRepositorio.asp';
        }

        function exibirLoading(exibir) {
            $('#loadingOverlay').toggleClass('d-none', !exibir).toggleClass('d-flex', exibir);
        }

        function navegarParaGrade() {
            window.location.href = 'SolicitacaoClienteExtensaoGrade.asp';
        }

        function navegarParaAba(numeroAba) {
            if (numeroAba < 1 || numeroAba > totalAbas) {
                return;
            }

            abaAtual = numeroAba;

            $('.tab-button').removeClass('active');
            $(`.tab-button[data-tab="${numeroAba}"]`).addClass('active');

            $('.tab-content').removeClass('active');
            $(`.tab-content[data-tab-content="${numeroAba}"]`).addClass('active');
        }

        function inicializarSelect2() {
            $('.select2').select2({
                theme: 'bootstrap-5',
                width: '100%',
                dropdownCssClass: 'z-index-9999'
            });
        }

        /* ---------- Normalizacao de propriedades do anexo ---------- */

        function obterStatusAnexo(anexo) {
            const status = (anexo.status || anexo.Status || STATUS_ANEXO.PENDENTE).toUpperCase();
            return STATUS_ANEXO[status] || STATUS_ANEXO.PENDENTE;
        }

        function obterMotivoAnexo(anexo) {
            return anexo.motivoInvalidacao || anexo.MotivoInvalidacao || '';
        }

        function obterIconeAnexo(nome) {
            const ext = (String(nome).split('.').pop() || '').toLowerCase();

            if (ext === 'pdf') {
                return { bi: 'bi-file-earmark-pdf', classe: 'ico-pdf' };
            }
            if (['jpg', 'jpeg', 'png', 'gif', 'webp', 'bmp'].includes(ext)) {
                return { bi: 'bi-file-earmark-image', classe: 'ico-img' };
            }
            if (['doc', 'docx', 'odt'].includes(ext)) {
                return { bi: 'bi-file-earmark-word', classe: 'ico-doc' };
            }
            if (['xls', 'xlsx', 'csv', 'ods'].includes(ext)) {
                return { bi: 'bi-file-earmark-spreadsheet', classe: 'ico-xls' };
            }
            return { bi: 'bi-file-earmark-text', classe: 'ico-generico' };
        }

        function escaparHtml(texto) {
            return $('<div>').text(texto == null ? '' : texto).html();
        }

        /* ---------- Carregamento ---------- */

        function carregarDadosCadastro() {
            return $.ajax({
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

                    preencherFormulario(response.dados);
                },
                error: function(xhr, status, error) {
                    console.error('Erro ao carregar dados:', error);
                    Toast.error('Erro ao carregar dados');
                }
            });
        }

          function preencherFormulario(dados) {
            if (!dados) {
                return;
            }

            dadosSolicitacao = dados;
            listaAnexos = dados.anexos || [];

            $('#codigoPlantaOrigem').val(dados.codigoPlantaOrigem || '').trigger('change.select2');
            $('#codigoPlantaDestino').val(dados.codigoPlantaDestino || '').trigger('change.select2');
            $('#codigoCliente').val(dados.codigoCliente || '');
            $('#codigoLoja').val(dados.codigoLoja || '');
            $('#razaoSocial').val(dados.razaoSocial || '');
            $('#cpfCnpj').val(dados.cpfCnpj || '').trigger('input');
            $('#codigoVendedor').val(dados.codigoVendedor || '');
            $('#nomeVendedor').val(dados.nomeVendedor || '');
            $('#tes').val(dados.tes || '');
            $('#toleranciaQuantidade').val(dados.toleranciaQuantidade || '');
            $('#toleranciaDias').val(dados.toleranciaDias || '');

            Promise.all([
                carregarNaturezas(dados.codigoPlantaOrigem)
            ]).then(function() {
                $('#natureza').val(dados.natureza).trigger('change.select2');
            });
            
            inicializarAvaliacoesAnexos();
            renderizarAnexos();
        }

        function carregarNaturezas(codigoPlanta) {
            const $select = $('#natureza');
            $select.empty().append('<option value="">Selecione um item</option>');

            return $.ajax({
                url: repoUrl(),
                type: 'GET',
                data: { acao: 'listarNaturezas', codigoPlanta: codigoPlanta },
                dataType: 'json',
                success: function(response) {
                    if (!response.sucesso) {
                        console.error('Erro ao carregar naturezas: ' + response.mensagem);
                        return;
                    }

                    (response.dados || []).forEach(function(item) {
                        $select.append($('<option>', { value: item.codigo, text: item.descricao }));
                    });
                },
                error: function(xhr, status, error) {
                    console.error('Erro ao carregar as naturezas:', error);
                }
            });
        }

        /* ---------- Estado de avaliacao dos anexos ---------- */

        function inicializarAvaliacoesAnexos() {
            avaliacoesAnexos = {};

            listaAnexos.forEach(function(anexo) {
                const anexoId = anexo.id;
                if (anexoId == null) {
                    return;
                }

                avaliacoesAnexos[anexoId] = {
                    status: obterStatusAnexo(anexo),
                    motivoInvalidacao: obterMotivoAnexo(anexo)
                };
            });
        }

        function renderizarAnexos() {
            const $container = $('#anexosListaContainer');

            if (!listaAnexos.length) {
                $('#avalResumo').addClass('d-none');
                $container.html(`
                    <div class="alert alert-info mb-0">
                        <i class="bi bi-info-circle me-1"></i>
                        Nenhum anexo ativo encontrado para esta solicitação.
                    </div>
                `);
                atualizarResumoAvaliacao();
                return;
            }

            const cards = listaAnexos.map(function(anexo) {
                const anexoId = anexo.id;
                const aval = avaliacoesAnexos[anexoId] || { status: STATUS_ANEXO.PENDENTE, motivoInvalidacao: '' };

                const nomeOriginal = anexo.nome;
                const nome = escaparHtml(nomeOriginal);
                const caminho = anexo.caminho;
                const icone = obterIconeAnexo(nomeOriginal);

                const ehInvalido = aval.status === STATUS_ANEXO.INVALIDO;
                const ehValido = aval.status === STATUS_ANEXO.VALIDO;
                const classeItem = ehInvalido ? 'is-invalido' : (ehValido ? 'is-valido' : 'is-pendente');

                return `
                    <div class="anexo-item ${classeItem}" data-anexo-id="${anexoId}">
                        <div class="anexo-main">
                            <a class="anexo-link" href="/Clientes_Docs/${caminho}" target="_blank" rel="noopener" title="Abrir documento">
                                <span class="anexo-icon-wrap ${icone.classe}">
                                    <i class="bi ${icone.bi}"></i>
                                </span>
                                <span class="anexo-nome">${nome}</span>
                                <i class="bi bi-box-arrow-up-right anexo-ext-ico"></i>
                            </a>

                            <div class="btn-group btn-group-sm anexo-toggle" role="group" aria-label="Status do documento">
                                <input type="radio" class="btn-check" name="statusAnexo-${anexoId}"
                                       id="anexoValido-${anexoId}" autocomplete="off" data-status="${STATUS_ANEXO.VALIDO}"
                                       ${ehValido ? 'checked' : ''}>
                                <label class="btn btn-outline-success" for="anexoValido-${anexoId}">
                                    <i class="bi bi-check-lg me-1"></i>Válido
                                </label>

                                <input type="radio" class="btn-check" name="statusAnexo-${anexoId}"
                                       id="anexoInvalido-${anexoId}" autocomplete="off" data-status="${STATUS_ANEXO.INVALIDO}"
                                       ${ehInvalido ? 'checked' : ''}>
                                <label class="btn btn-outline-danger" for="anexoInvalido-${anexoId}">
                                    <i class="bi bi-x-lg me-1"></i>Inválido
                                </label>
                            </div>
                        </div>

                        <div class="anexo-motivo ${ehInvalido ? '' : 'd-none'}" data-anexo-motivo>
                            <label class="form-label mb-1">
                                Motivo da invalidação <span class="required-field">*</span>
                            </label>
                            <textarea class="form-control form-control-sm anexo-motivo-input"
                                      rows="2" maxlength="500"
                                      placeholder="Descreva por que este documento é inválido...">${escaparHtml(aval.motivoInvalidacao)}</textarea>
                        </div>
                    </div>
                `;
            }).join('');

            $container.html(`<div class="anexos-list">${cards}</div>`);
            atualizarResumoAvaliacao();
        }

        function contarStatus() {
            const valores = Object.values(avaliacoesAnexos);
            return {
                total: valores.length,
                validos: valores.filter(a => a.status === STATUS_ANEXO.VALIDO).length,
                invalidos: valores.filter(a => a.status === STATUS_ANEXO.INVALIDO).length,
                pendentes: valores.filter(a => a.status === STATUS_ANEXO.PENDENTE).length
            };
        }

        function atualizarResumoAvaliacao() {
            const c = contarStatus();

            const $resumo = $('#avalResumo');
            if (c.total > 0) {
                $resumo.removeClass('d-none').html(`
                    <span class="resumo-chip"><i class="bi bi-files"></i> ${c.total} ${c.total === 1 ? 'documento' : 'documentos'}</span>
                    <span class="resumo-chip chip-valido"><i class="bi bi-check-circle"></i> ${c.validos} válidos</span>
                    <span class="resumo-chip chip-invalido"><i class="bi bi-x-circle"></i> ${c.invalidos} inválidos</span>
                    <span class="resumo-chip chip-pendente"><i class="bi bi-hourglass-split"></i> ${c.pendentes} pendentes</span>
                `);
            } else {
                $resumo.addClass('d-none');
            }

            const $meta = $('#menuContadorAnexos').removeClass('meta-pendente meta-invalido meta-ok');
            if (!c.total) {
                $meta.text('0');
            } else if (c.pendentes > 0) {
                $meta.addClass('meta-pendente').text(`${c.pendentes} pend.`);
            } else if (c.invalidos > 0) {
                $meta.addClass('meta-invalido').text(`${c.invalidos} inv.`);
            } else {
                $meta.addClass('meta-ok').text('OK');
            }

        }

        /* ---------- Eventos da avaliacao por anexo (delegados) ---------- */

        function anexarEventosAnexos() {
            $('#anexosListaContainer').on('change', '.btn-check', function() {
                const $item = $(this).closest('.anexo-item');
                const anexoId = $item.data('anexo-id');
                const status = $(this).data('status');

                if (!avaliacoesAnexos[anexoId]) {
                    avaliacoesAnexos[anexoId] = { status: STATUS_ANEXO.PENDENTE, motivoInvalidacao: '' };
                }

                avaliacoesAnexos[anexoId].status = status;

                const $motivo = $item.find('[data-anexo-motivo]');
                const $motivoInput = $item.find('.anexo-motivo-input');

                if (status === STATUS_ANEXO.INVALIDO) {
                    $motivo.removeClass('d-none');
                    $item.removeClass('is-valido is-pendente').addClass('is-invalido');
                    $motivoInput.trigger('focus');
                } else {
                    $motivo.addClass('d-none');
                    $motivoInput.val('').removeClass('just-validate-error-field');
                    $item.find('.erro-validacao').remove();
                    avaliacoesAnexos[anexoId].motivoInvalidacao = '';
                    $item.removeClass('is-invalido is-pendente').addClass('is-valido');
                }

                atualizarResumoAvaliacao();
            });

            $('#anexosListaContainer').on('input', '.anexo-motivo-input', function() {
                const anexoId = $(this).closest('.anexo-item').data('anexo-id');

                if (!avaliacoesAnexos[anexoId]) {
                    return;
                }

                avaliacoesAnexos[anexoId].motivoInvalidacao = ($(this).val() || '').trim();

                if (avaliacoesAnexos[anexoId].motivoInvalidacao) {
                    $(this).removeClass('just-validate-error-field')
                           .closest('[data-anexo-motivo]').find('.erro-validacao').remove();
                }
            });
        }

        /* ---------- Validacao ---------- */

        function limparErrosValidacao() {
            $('.just-validate-error-label, .erro-validacao').remove();
            $('.just-validate-error-field').removeClass('just-validate-error-field');
        }

        function rolarAteItem(anexoId) {
            $(`.anexo-item[data-anexo-id="${anexoId}"]`)[0]
                ?.scrollIntoView({ behavior: 'smooth', block: 'center' });
        }

        function validarAvaliacao(flagAprovar) {
            limparErrosValidacao();

            const observacao = ($('#observacaoAvaliacao').val() || '').trim();
            const entradas = Object.entries(avaliacoesAnexos);

            if (!entradas.length) {
                Toast.warning('Nenhum anexo para avaliar nesta solicitação.');
                return false;
            }

            const pendentes = entradas.filter(([, a]) => a.status === STATUS_ANEXO.PENDENTE);
            if (pendentes.length > 0) {
                Toast.warning(`Avalie todos os documentos antes de concluir. Faltam ${pendentes.length}.`);
                navegarParaAba(1);
                rolarAteItem(pendentes[0][0]);
                return false;
            }

            const invalidosSemMotivo = entradas.filter(([, a]) =>
                a.status === STATUS_ANEXO.INVALIDO && !a.motivoInvalidacao
            );

            if (invalidosSemMotivo.length > 0) {
                invalidosSemMotivo.forEach(([anexoId]) => {
                    const $item = $(`.anexo-item[data-anexo-id="${anexoId}"]`);
                    const $input = $item.find('.anexo-motivo-input').addClass('just-validate-error-field');
                    if (!$item.find('.erro-validacao').length) {
                        $input.after('<div class="erro-validacao">Informe o motivo da invalidação.</div>');
                    }
                });

                Toast.warning('Informe o motivo de cada documento marcado como inválido.');
                navegarParaAba(1);
                rolarAteItem(invalidosSemMotivo[0][0]);
                return false;
            }

            const possuiInvalido = entradas.some(([, a]) => a.status === STATUS_ANEXO.INVALIDO);
            if (flagAprovar && possuiInvalido) {
                Toast.warning('Existe documento inválido. A solicitação só pode ser devolvida.');
                return false;
            }

            if (!flagAprovar && !observacao) {
                navegarParaAba(1);
                $('#observacaoAvaliacao')
                    .addClass('just-validate-error-field')
                    .after('<div class="just-validate-error-label">Observação obrigatória para devolução.</div>');
                Toast.warning('Informe o motivo da devolução.');
                $('#observacaoAvaliacao')[0]?.scrollIntoView({ behavior: 'smooth', block: 'center' });
                $('#observacaoAvaliacao').trigger('focus');
                return false;
            }

            return true;
        }

        /* ---------- Envio ---------- */

        function montarAnexosPayload() {
            return Object.entries(avaliacoesAnexos).map(([anexoId, aval]) => {
                const ehInvalido = aval.status === STATUS_ANEXO.INVALIDO;
                return {
                    id: anexoId,
                    status: aval.status,
                    flagInvalido: ehInvalido ? 1 : 0,
                    motivoInvalidacao: ehInvalido ? (aval.motivoInvalidacao || '').replaceAll("'", '"') : ''
                };
            });
        }

        function avaliarSolicitacao(flagAprovar) {
            if (!validarAvaliacao(flagAprovar)) {
                return;
            }

            const observacao = ($('#observacaoAvaliacao').val() || '').trim().replaceAll("'", '"');
            const anexosPayload = montarAnexosPayload();

            exibirLoading(true);

            $.ajax({
                url: repoUrl(),
                type: 'POST',
                dataType: 'json',
                data: {
                    acao: 'avaliacaoCDAG',
                    id: id,
                    flagAprovar: flagAprovar ? 1 : 0,
                    observacao: observacao,
                    anexos: JSON.stringify(anexosPayload)
                },
                success: function(response) {
                    if (!response.sucesso) {
                        Toast.error(response.mensagem);
                        return;
                    }

                    Toast.success(response.mensagem || (
                        flagAprovar
                            ? 'Solicitação aprovada com sucesso!'
                            : 'Solicitação devolvida com sucesso!'
                    ));

                    setTimeout(navegarParaGrade, 500);
                },
                error: function(xhr, status, error) {
                    console.error('Erro ao avaliar solicitação:', error);
                    Toast.error('Erro ao avaliar solicitação');
                },
                complete: function() {
                    exibirLoading(false);
                }
            });
        }

        /* ---------- Eventos gerais ---------- */

        function anexarEventos() {
            $('#formularioAvaliacao').on('submit', function(e) {
                e.preventDefault();
                return false;
            });

            $('.tab-button').on('click', function() {
                const tabNumber = parseInt($(this).data('tab'));
                if (tabNumber === abaAtual) {
                    return;
                }
                navegarParaAba(tabNumber);
            });

            $('#btnVoltarGrade').on('click', navegarParaGrade);

            $('#btnAprovar').on('click', function() {
                avaliarSolicitacao(true);
            });

            $('#btnDevolver').on('click', function() {
                avaliarSolicitacao(false);
            });

            anexarEventosAnexos();
        }

        $(document).ready(async function() {
            const urlParams = new URLSearchParams(window.location.search);
            id = urlParams.get('id');
            acao = (urlParams.get('acao') || '').toUpperCase();

            anexarEventos();
            navegarParaAba(1);
            inicializarSelect2();

            if (!id) {
                $('#titulo').text('Avaliação de Solicitação');
                Toast.warning('Informe uma solicitação para carregar os dados.');
                return;
            }

            $('#titulo').text(`Avaliação de Solicitação #${id}`);
            await carregarDadosCadastro();
        });
    </script>
</head>
<body>
    <div class="container">
        <div class="d-flex my-4 justify-content-between align-items-center flex-wrap gap-2">
            <div class="d-flex align-items-center">
                <button type="button" class="btn btn-sm btn-light me-2 btn-back" id="btnVoltarGrade" title="Voltar para a lista">
                    <i class="bi bi-arrow-left"></i>
                </button>

                <h3 id="titulo" class="mb-0"></h3>
            </div>
        </div>

        <div id="loadingOverlay" class="loading-overlay d-none">
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
                            <span class="tab-title">Avaliação</span>
                            <span class="tab-meta" id="menuContadorAnexos">0</span>
                        </button>

                        <div class="tab-group-title">Consulta</div>

                        <button type="button" class="tab-button" data-tab="2">
                            <i class="tab-icon bi bi-clipboard-check"></i>
                            <span class="tab-title">Dados Gerais</span>
                        </button>
                    </div>
                </div>

                <div class="tab-content-wrapper">
                    <form id="formularioAvaliacao" name="formularioAvaliacao" class="form-content" novalidate>
                        <!-- Aba 1: Avaliacao (anexos + parecer) -->
                        <div class="tab-content active" data-tab-content="1">
                            <div class="section-header">
                                <div class="section-title">Avaliação da documentação</div>
                                <p class="section-subtitle">
                                    Marque cada documento como válido ou inválido. Para inválidos, informe o motivo. Em seguida, registre o parecer e conclua.
                                </p>
                            </div>

                            <div id="avalResumo" class="aval-resumo d-none"></div>

                            <div id="anexosListaContainer"></div>

                            <div class="bloco-parecer">
                                <label for="observacaoAvaliacao" class="form-label">
                                    Parecer geral
                                    <span class="text-muted fw-normal">(obrigatório para devolução)</span>
                                </label>

                                <textarea
                                    class="form-control form-control-sm"
                                    id="observacaoAvaliacao"
                                    rows="4"
                                    maxlength="1000"
                                    placeholder="Resumo da avaliação para o solicitante..."
                                ></textarea>

                                <div class="readonly-alert mt-3">
                                    <i class="bi bi-info-circle me-1"></i>
                                    Documentos marcados como inválidos exigem um motivo e impedem a aprovação da solicitação.
                                </div>
                            </div>
                        </div>

                        <!-- Aba 2: Dados Gerais (consulta) -->
                        <div class="tab-content" data-tab-content="2">
                            <div class="section-header">
                                <div class="section-title">Dados da Solicitação</div>
                                <p class="section-subtitle">Informações básicas da solicitação selecionada em modo consulta.</p>
                            </div>

                            <div class="row g-3">
                                <div class="col-md-6">
                                    <label for="codigoPlantaOrigem" class="form-label">Planta Origem <span class="required-field">*</span></label>
                                    <select disabled class="form-select form-select-sm select2" name="codigoPlantaOrigem" id="codigoPlantaOrigem">
                                        <option value="">Selecione um item</option>
                                        <option value="00">PIRAPETINGA</option>
                                        <option value="04">UBERABA</option>
                                        <option value="05">BENTO GONÇALVES</option>
                                        <option value="06">MARANGUAPE</option>
                                        <option value="07">SAQUAREMA</option>
                                    </select>
                                </div>

                                <div class="col-md-6">
                                    <label for="codigoPlantaDestino" class="form-label">Planta Destino <span class="required-field">*</span></label>
                                    <select disabled class="form-select form-select-sm select2" name="codigoPlantaDestino" id="codigoPlantaDestino">
                                        <option value="">Selecione um item</option>
                                        <option value="00">PIRAPETINGA</option>
                                        <option value="04">UBERABA</option>
                                        <option value="05">BENTO GONÇALVES</option>
                                        <option value="06">MARANGUAPE</option>
                                        <option value="07">SAQUAREMA</option>
                                    </select>
                                </div>

                                <div class="col-12 col-md-3">
                                    <label for="codigoCliente" class="form-label">Cód. Cliente <span class="required-field me-2">*</span></label>
                                    <input disabled onblur="obterCliente()" type="text" class="form-control form-control-sm" id="codigoCliente" name="codigoCliente">
                                </div>
                                
                                <div class="col-12 col-md-2">
                                    <label for="codigoLoja" class="form-label">Cód. Loja <span class="required-field">*</span></label>
                                    <input disabled type="text" onblur="obterCliente()" class="form-control form-control-sm" id="codigoLoja" name="codigoLoja">
                                </div>
                                
                                <div class="col-12 col-md-4">
                                    <label for="razaoSocial" class="form-label">Razão Social <span class="required-field">*</span></label>
                                    <input disabled type="text" class="form-control form-control-sm" name="razaoSocial" id="razaoSocial" maxlength="160" />
                                </div>

                                <div class="col-12 col-md-3">
                                    <label for="cpfCnpj" class="form-label" id="labelCpfCnpj">CNPJ/CPF <span class="required-field">*</span></label>
                                    <input disabled type="text" class="form-control form-control-sm" name="cpfCnpj" id="cpfCnpj" />
                                </div>
                                
                                <div class="col-12 col-md-4">
                                    <label for="codigoVendedor" class="form-label">Cód. Vendedor <span class="required-field me-2">*</span></label>
                                    <input disabled type="text" class="form-control form-control-sm" id="codigoVendedor" name="codigoVendedor">
                                </div>

                                <div class="col-12 col-md-4">
                                    <label for="nomeVendedor" class="form-label">Nome do Vendedor <span class="required-field">*</span></label>
                                    <input disabled type="text" class="form-control form-control-sm" id="nomeVendedor" name="nomeVendedor">
                                </div>

                                <div class="col-md-4">
                                    <label for="tes" class="form-label">TES <span class="required-field">*</span></label>
                                    <input disabled type="text" class="form-control form-control-sm" name="tes" id="tes" maxlength="20" />
                                </div>

                                <div class="col-md-4">
                                    <label for="natureza" class="form-label">Natureza <span class="required-field">*</span></label>
                                    <select disabled class="form-select form-select-sm select2" name="natureza" id="natureza"></select>
                                </div>

                                <div class="col-md-4">
                                    <label for="toleranciaQuantidade" class="form-label">Tolerância de Quantidade +/- <span class="required-field">*</span></label>
                                    <input disabled type="text" class="form-control form-control-sm percentual" name="toleranciaQuantidade" id="toleranciaQuantidade" />
                                </div>

                                <div class="col-md-4">
                                    <label for="toleranciaDias" class="form-label">Tolerância de Dias +/- <span class="required-field">*</span></label>
                                    <input disabled type="text" class="form-control form-control-sm" name="toleranciaDias" id="toleranciaDias" />
                                </div>
                            </div>
                        </div>
                    </form>

                    <div class="action-bar">
                        <div class="action-buttons">
                            <button type="button" class="btn btn-sm btn-danger" id="btnDevolver">
                                <i class="bi bi-arrow-counterclockwise me-1"></i>Devolver
                            </button>

                            <button type="button" class="btn btn-sm btn-success" id="btnAprovar">
                                <i class="bi bi-check-lg me-1"></i>Aprovar
                            </button>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    </div>
</body>
</html>
