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

        .form-switch {
            display: flex;
            align-items: end;
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

        th,
        td {
            white-space: nowrap;
            vertical-align: top;
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
    <script src="https://cdn.jsdelivr.net/npm/just-validate@latest/dist/just-validate.production.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/jquery-maskmoney/3.0.2/jquery.maskMoney.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/jquery.mask/1.14.15/jquery.mask.min.js"></script>

    <script src="/corporativo/components/toast.js?v=<%=Now()%>"></script>
    <script src="/corporativo/components/confirmation.js?v=<%=Now()%>"></script>
    <script src="/corporativo/components/upload/fileUpload.js?v=<%=Now()%>"></script>
    <script src="/corporativo/components/upload/documentUpload.js?v=<%=Now()%>"></script>
    <script src="/corporativo/funcoes/funcoesJAVASCRIPT.js?v=<%=Now()%>"></script>

    <script>
        let id = 0;
        let preCadastro = '';
        let abaAtual = 1;
        let totalAbas = 3;
        let abasValidadas = new Set();
        let validacoes = {};
        let acao = 'CRIAR';

        // Busca de Cliente
        let clienteSelecionado = null;
        let resultadosCliente = [];

        // Anexos
        let listaAnexos = [];

        function exibirLoading(exibir) {
            if (exibir) {
                $('#loadingOverlay').removeClass('d-none').addClass('d-flex');
            } else {
                $('#loadingOverlay').removeClass('d-flex').addClass('d-none');
            }
        }

        function formularioSomenteVisualizacao() {
            return acao === 'VISUALIZAR';
        }
        function configurarModoFormulario() {
            if (!formularioSomenteVisualizacao()) return;

            $('#divAcoes').addClass('d-none');
            $('#formularioExtensao').find('input, select, textarea').prop('disabled', true);
            $('#formularioExtensao').find('.select2').trigger('change.select2');

            // Em modo visualização, o componente continua visível para permitir abrir o anexo.
            // Apenas bloqueamos ações de anexar/trocar/remover.
            $('document-upload').each(function() {
                if (typeof this.disable === 'function') {
                    this.disable();
                } else {
                    $(this).attr('disabled', 'disabled');
                }
            });
        }

        function inicializarSelect2() {
            $('.select2').select2({
                theme: 'bootstrap-5',
                width: '100%',
                dropdownCssClass: 'z-index-9999'
            });
        }

        function aplicarRegrasDocumentacao() {
            const tipoPessoa = $('#hidTipoPessoa').val();
            const possuiIe = $('#hidInscricaoEstadual').val() != '';
            const possuiIcmsDiferido = $('#icmsDiferido').val() === '1';

            $('.documentos-pj, .documentos-pf').addClass('d-none');
            $('.documento-regra').addClass('d-none');

            if (tipoPessoa === 'J') {
                $('.documentos-pj').removeClass('d-none');
                $('.documento-pj-obrigatorio').removeClass('d-none');

                if (possuiIe) {
                    $('.documento-pj-inscricao-estadual').removeClass('d-none');
                } else if ($('#contribuinteIcms').val() !== '') {
                    $('.documento-pj-carta-isencao').removeClass('d-none');
                }

                if (possuiIcmsDiferido) {
                    $('.documento-pj-deferimento').removeClass('d-none');
                }
            }

            if (tipoPessoa === 'F') {
                $('.documentos-pf').removeClass('d-none');
                $('.documento-pf-obrigatorio').removeClass('d-none');

                if (possuiIe) {
                    $('.documento-pf-inscricao-estadual').removeClass('d-none');
                } else if ($('#contribuinteIcms').val() !== '') {
                    $('.documento-pf-carta-isencao').removeClass('d-none');
                }

                if (possuiIcmsDiferido) {
                    $('.documento-pf-deferimento').removeClass('d-none');
                }
            }
        }

        function aplicarCamposCondicionais() {
            const tipoPessoa = $('#hidTipoPessoa').val();

            if (tipoPessoa === 'J') {
                $('#cpfCnpj').mask('AA.AAA.AAA/AAAA-00', {
                    translation: { 'A': { pattern: /[A-Za-z0-9]/ } },
                    onKeyPress: function(value, event, field) {
                        field.val(value.toUpperCase());
                    }
                });
                
                $('#cpfCnpj').off('input.apenasMaiusculo').on('input.apenasMaiusculo', function() {
                    this.value = this.value.toUpperCase();
                });

                $('#labelCpfCnpj').html('CNPJ <span class="required-field">*</span>');
            }

            if (tipoPessoa === 'F') {
                $('#cpfCnpj').mask('000.000.000-00');
                $('#labelCpfCnpj').html('CPF <span class="required-field">*</span>');
            }

            aplicarRegrasDocumentacao();
            carregarAnexosExistentesNosUploads();
            atualizarContadorAnexos();
        }

        function normalizarTextoComparacao(valor) {
            return String(valor || '').trim().toLowerCase();
        }

        function obterNomeDocumentoAnexo(anexo) {
            return anexo.nome || '';
        }

        function obterIdDocumentoAnexo(anexo) {
            return anexo.id || '';
        }

        function localizarAnexoPorNome(nomeDocumento) {
            const nomeNormalizado = normalizarTextoComparacao(nomeDocumento);

            return (listaAnexos || []).find(anexo =>
                normalizarTextoComparacao(obterNomeDocumentoAnexo(anexo)) === nomeNormalizado
            );
        }

        function montarItemAnexoParaUpload(anexo) {
            if (!anexo) return null;

            const caminho = anexo.caminho || '';

            return {
                id: obterIdDocumentoAnexo(anexo),
                nome: obterNomeDocumentoAnexo(anexo),
                caminho: caminho,
                dataInclusao: anexo.dataInclusao || '',
                invalidated: parseInt(anexo.flagInvalido) || false,
                invalidationReason: anexo.motivoInvalido || '',
                url: caminho ? `/Clientes_Docs/${caminho}` : ''
            };
        }

        function carregarAnexosExistentesNosUploads() {
            $('document-upload').each(function() {
                const componente = this;
                const nomeDocumento = componente.getAttribute('name');
                const anexo = localizarAnexoPorNome(nomeDocumento);

                if (typeof componente.isDirty === 'function' && componente.isDirty()) {
                    return;
                }

                if (anexo && typeof componente.loadExistingFile === 'function') {
                    componente.loadExistingFile(montarItemAnexoParaUpload(anexo));
                }
            });
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

        function renderizarAnexos() {
            carregarAnexosExistentesNosUploads();
            atualizarContadorAnexos();
        }

        function atualizarContadorAnexos() {
            let total = 0;

            $('document-upload').each(function() {
                if (!documentoUploadEstaDisponivel(this)) return;

                if (typeof this.hasFile === 'function' && this.hasFile()) {
                    total++;
                }
            });

            $('#menuContadorAnexos').text(total);
        }

        function documentoUploadEstaDisponivel(componente) {
            const $componente = $(componente);

            if ($componente.attr('disabled') && !formularioSomenteVisualizacao()) {
                return false;
            }

            return $componente.parents().addBack().filter(function() {
                const $elemento = $(this);

                if ($elemento.hasClass('tab-content')) {
                    return false;
                }

                return $elemento.hasClass('d-none') || this.style.display === 'none';
            }).length === 0;
        }

        function coletarDocumentos() {
            const componentes = $('document-upload').filter(function() {
                return documentoUploadEstaDisponivel(this);
            }).toArray();
            const documentos = [];
            let tudoOk = true;

            componentes.forEach(c => {
                const nomeDocumento = c.getAttribute('name');

                if (typeof c.validate === 'function' && !c.validate()) {
                    tudoOk = false;
                    return;
                }

                if (typeof c.getValue !== 'function') {
                    return;
                }

                const valor = c.getValue();

                if (valor.action === 'added' || valor.action === 'replaced') {
                    valor.name = nomeDocumento;
                }

                documentos.push(valor);
            });

            if (!tudoOk) {
                return null;
            }

            return documentos;
        }

        function inicializarValidacoes() {
            const obrigatorio = (msg) => ({
                validator: (value, fields) => value !== null && value !== undefined && String(value).trim() !== '',
                errorMessage: msg
            });

            const somenteNumerosMaximoDigitosQuandoInformado = (maximoDigitos, msg) => ({
                validator: (value, fields) => {
                    if (value === null || value === undefined || String(value).trim() === '') return true;
                    const valor = String(value).trim();
                    return /^\d+$/.test(valor) && valor.length <= maximoDigitos;
                },
                errorMessage: msg
            });

            validacoes.aba1 = new JustValidate('#formularioExtensao', { lockForm: false });
            validacoes.aba1
                .addField('#codigoPlantaOrigem', [obrigatorio('Planta origem é obrigatória')])
                .addField('#codigoPlantaDestino', [obrigatorio('Planta destino é obrigatória')])
                .addField('#codigoCliente', [obrigatorio('Código do cliente é obrigatório')])
                .addField('#codigoLoja', [obrigatorio('Loja do cliente é obrigatória')])
                .addField('#razaoSocial', [obrigatorio('Razão Social é obrigatória')])
                .addField('#cpfCnpj', [obrigatorio('CNPJ/CPF é obrigatório')])
                .addField('#codigoVendedor', [obrigatorio('Vendedor é obrigatório')])
                .addField('#natureza', [obrigatorio('Natureza é obrigatória')])
                .addField('#toleranciaQuantidade', [obrigatorio('Tolerância de quantidade é obrigatória')])
                .addField('#toleranciaDias', [obrigatorio('Tolerância de dias é obrigatória')]);

            validacoes.aba2 = new JustValidate('#formularioExtensao', { lockForm: false });
            validacoes.aba2
                .addField('#icmsDiferido', [obrigatorio('ICMS diferido é obrigatório')])
                .addField('#contribuinteIcms', [obrigatorio('Contribuinte de ICMS é obrigatório')])
                .addField('#tes', [somenteNumerosMaximoDigitosQuandoInformado(3, 'TES deve conter apenas números e no máximo 3 dígitos')]);
        }

        function navegarParaGrade() {
            window.location.href = 'SolicitacaoClienteExtensaoGrade.asp';
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

        function carregarDadosCadastro() {
            const params = {
                acao: 'obterPorId',
                id: id
            };

            return $.ajax({
                url: 'SolicitacaoClienteExtensaoRepositorio.asp',
                type: 'GET',
                data: params,
                dataType: 'json',
                success: function(response) {
                    if (!response.sucesso) {
                        Toast.error(response.mensagem);
                        return;
                    }

                    preencherFormularioEdicao(response.dados);
                },
                error: function(xhr, status, error) {
                    console.error('Erro ao carregar dados:', error);
                    Toast.error('Erro ao carregar dados');
                }
            });
        }

        function carregarNaturezas(codigoPlanta) {
            const $select = $('#natureza');
            $select.empty().append('<option value="">Selecione um item</option>');

            return $.ajax({
                url: 'SolicitacaoClienteExtensaoRepositorio.asp?v=<%=Now()%>',
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

        function carregarTiposContribuinte() {
            const $select = $('#contribuinteIcms');
            $select.empty().append('<option value="">Selecione um item</option>');

            return $.ajax({
                url: 'SolicitacaoClienteExtensaoRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: { acao: 'listarTiposContribuinte' },
                dataType: 'json',
                success: function(response) {
                    if (!response.sucesso) {
                        console.error('Erro carregar tipos de contribuinte: ' + response.mensagem);
                        return;
                    }

                    (response.dados || []).forEach(function(item) {
                        $select.append($('<option>', { value: item.codigo, text: item.descricao }));
                    });
                    $select.trigger('change.select2');
                },
                error: function(xhr, status, error) {
                    console.error('Erro ao carregar tipos de contribuinte:', error);
                }
            });
        }

        function preencherFormularioEdicao(dados) {
            if (!dados) return;

            // Desabilita campos não editáveis
            $('#btnBuscarCliente, #codigoPlantaOrigem, #codigoPlantaDestino').prop('disabled', true);

            // Preenche formulário
            $('#hidTipoPessoa').val(dados.tipoPessoa || '');
            $('#hidInscricaoEstadual').val(dados.inscricaoEstadual || '');
            $('#codigoPlantaOrigem').val(dados.codigoPlantaOrigem || '').trigger('change.select2');
            $('#codigoPlantaDestino').val(dados.codigoPlantaDestino || '').trigger('change.select2');
            $('#codigoCliente').val(dados.codigoCliente || '');
            $('#codigoLoja').val(dados.codigoLoja || '');
            $('#razaoSocial').val(dados.razaoSocial || '');
            $('#cpfCnpj').val(dados.cpfCnpj || '').trigger('input');
            $('#codigoVendedor').val(dados.codigoVendedor || '');
            $('#nomeVendedor').val(dados.nomeVendedor || '');
            $('#natureza').val(dados.natureza || '');
            $('#toleranciaQuantidade').val(dados.toleranciaQuantidade || '');
            $('#toleranciaDias').val(dados.toleranciaDias || '');

            $('#contribuinteIcms').val(dados.contribuinteIcms).trigger('change.select2');
            $('#icmsDiferido').val(dados.icmsDiferido).trigger('change.select2');
            $('#tes').val(dados.tes);

            Promise.all([
                carregarNaturezas(dados.codigoPlantaOrigem)
            ]).then(function() {
                $('#natureza').val(dados.natureza).trigger('change.select2');
            });

            listaAnexos = dados.anexos || [];

            aplicarCamposCondicionais();
            renderizarAnexos();
            configurarModoFormulario();
        }

        function irProximo() {
            navegarParaAba(abaAtual + 1);
        }

        function irAnterior() {
            navegarParaAba(abaAtual - 1);
        }
        
        function anexarEventos() {
            $('#codigoPlantaOrigem').on('change', function() {
                carregarNaturezas($(this).val());
            });

            $('#hidTipoPessoa').on('change', function() {
                $('#cpfCnpj').val('');
                aplicarCamposCondicionais();
            });

            $('#contribuinteIcms, #icmsDiferido').on('change', function() {
                aplicarCamposCondicionais();
            });

            $('#codigoPlantaOrigem, #codigoPlantaDestino').on('change', function() {
                if ($('#codigoPlantaOrigem').val() && $('#codigoPlantaOrigem').val() === $('#codigoPlantaDestino').val()) {
                    $('#codigoPlantaDestino').val('').trigger('change');
                    Toast.warning('A planta destino deve ser diferente da origem.');
                }
            });

            $(document).on('change upload-success delete-success', 'document-upload', function() {
                atualizarContadorAnexos();
            });

            /* Modal de buscar cliente */
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

            $('#btnSalvar').on('click', function() {
                if (!formularioSomenteVisualizacao()) {
                    validarFormulario();
                }
            });

            $('#formularioExtensao').on('submit', function(e) {
                e.preventDefault();
                if (!formularioSomenteVisualizacao()) {
                    validarFormulario();
                }
            });

            $('.tab-button').on('click', function() {
                const tabNumber = parseInt($(this).data('tab'));
                if (tabNumber === abaAtual) return;
                navegarParaAba(tabNumber);
            });
        }

        async function validarFormulario() {
            if (formularioSomenteVisualizacao()) return;

            const abas = [1, 2];
            const resultados = await Promise.all([
                validacoes.aba1.revalidate(),
                validacoes.aba2.revalidate(),
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

            if (!resultados.every(Boolean)) {
                const primeiraInvalida = abas[resultados.findIndex(v => !v)];
                navegarParaAba(primeiraInvalida);
                Toast.error('Existem campos obrigatórios pendentes. Verifique as seções destacadas.');
                return;
            }

            const documentos = coletarDocumentos();
            if (documentos === null) {
                marcarAbaComoInvalida(3);
                navegarParaAba(3);
                Toast.error('Existem documentos obrigatórios pendentes. Verifique a seção Anexos.');
                return;
            }

            marcarAbaComoCompleta(3);
            enviarFormulario(documentos);
        }

        function enviarFormulario(documentos) {
            if (formularioSomenteVisualizacao()) return;

            const dadosFormulario = {
                codigoPlantaOrigem: $('#codigoPlantaOrigem').val(),
                codigoPlantaDestino: $('#codigoPlantaDestino').val(),
                tipoPessoa: $('#hidTipoPessoa').val(),
                inscricaoEstadual: $('#hidInscricaoEstadual').val(),
                codigoCliente: $('#codigoCliente').val(),
                codigoLoja: $('#codigoLoja').val(),
                razaoSocial: $('#razaoSocial').val(),
                cpfCnpj: $('#cpfCnpj').val(),
                codigoVendedor: $('#codigoVendedor').val(),
                nomeVendedor: $('#nomeVendedor').val(),
                tes: $('#tes').val(),
                contribuinteIcms: $('#contribuinteIcms').val(),
                icmsDiferido: $('#icmsDiferido').val(),
                natureza: $('#natureza').val(),
                toleranciaQuantidade: $('#toleranciaQuantidade').val(),
                toleranciaDias: $('#toleranciaDias').val(),
                anexos: JSON.stringify(documentos)
            };

            if (id) {
                dadosFormulario.id = id;
                dadosFormulario.acao = 'editar';
            } else {
                dadosFormulario.acao = 'criar';
            }

            exibirLoading(true);

            $.ajax({
                url: 'SolicitacaoClienteExtensaoRepositorio.asp',
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

        function abrirModalCliente() {
            const codigoPlantaOrigem = $('#codigoPlantaOrigem').val();

            if (!codigoPlantaOrigem) { 
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
            const codigoPlantaOrigem = $('#codigoPlantaOrigem').val();

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
                url: 'SolicitacaoClienteExtensaoRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: {
                    acao: 'listarClientes',
                    codigoPlanta: codigoPlantaOrigem,
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
            aplicarCamposCondicionais();
        }

        
        function preencherDadosCliente(dados) {
            $('#razaoSocial').val(dados.razaoSocial || '');
            $('#codigoCliente').val(dados.codigoCliente || '');
            $('#codigoLoja').val(dados.codigoLoja || '');
            $('#cpfCnpj').val(dados.cpfCnpj || '').trigger('input');
            $('#codigoVendedor').val(dados.codigoVendedor || '');
            $('#nomeVendedor').val(dados.nomeVendedor || '');
            $('#municipio').val(dados.municipio || '');
            $('#hidTipoPessoa').val(dados.tipoPessoa);
            $('#hidInscricaoEstadual').val(dados.inscricaoEstadual || '');
        }

        $(document).ready(async function() {
            const urlParams = new URLSearchParams(window.location.search);
            id = urlParams.get('id');
            acao = urlParams.get('acao');

            await carregarTiposContribuinte();
       
            anexarEventos();
            navegarParaAba(1);
            inicializarSelect2();

            if (id) {
                await carregarDadosCadastro();

                if (formularioSomenteVisualizacao()) {
                    $('#titulo').text(`Visualizar Extensão #${id}`);
                } else {
                    $('#titulo').text(`Editar Extensão #${id}`);
                }
            } else {
                $('#titulo').text('Extensão de Cliente');
            }

            inicializarValidacoes();
            aplicarCamposCondicionais();
            configurarModoFormulario();
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
                            <span class="tab-title">Dados Gerais</span>
                        </button>

                        <div class="tab-group-title">Cadastro</div>

                        <button type="button" class="tab-button" data-tab="2">
                            <i class="tab-icon bi bi-receipt"></i>
                            <span class="tab-title">Fiscal</span>
                        </button>

                        <div class="tab-group-title">Adicionais</div>

                        <button type="button" class="tab-button" data-tab="3">
                            <i class="tab-icon bi bi-paperclip"></i>
                            <span class="tab-title">Anexos</span>
                            <span class="tab-meta" id="menuContadorAnexos">0</span>
                        </button>
                    </div>
                </div>

                <div class="tab-content-wrapper">
                    <form style="overflow:hidden;" id="formularioExtensao" name="formularioExtensao" novalidate>  
                        <input type="hidden" name="hidTipoPessoa" id="hidTipoPessoa" value="" />
                        <input type="hidden" name="hidInscricaoEstadual" id="hidInscricaoEstadual" value="" />

                        <div class="tab-content active" data-tab-content="1">
                            <div class="section-header">
                                <div class="section-title">Dados Solicitação</div>
                                <p class="section-subtitle">Informações de extensão de cadastro entre plantas.</p>
                            </div>

                            <div class="row g-3">
                                <div class="col-md-6">
                                    <label for="codigoPlantaOrigem" class="form-label">Planta Origem <span class="required-field">*</span></label>
                                    <select class="form-select form-select-sm select2" name="codigoPlantaOrigem" id="codigoPlantaOrigem">
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
                                    <select class="form-select form-select-sm select2" name="codigoPlantaDestino" id="codigoPlantaDestino">
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
                                    <div class="input-group">
                                        <button type="button" class="btn btn-sm btn-secondary" id="btnBuscarCliente" onclick="abrirModalCliente()">
                                            Buscar
                                            <i class="bi bi-box-arrow-up-right"></i>
                                        </button>
                                        <input disabled onblur="obterCliente()" type="text" class="form-control form-control-sm" id="codigoCliente" name="codigoCliente">
                                    </div>
                                    <small class="text-muted">Clique no botão para buscar ou digite um cód.</small>
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
                                    <label for="natureza" class="form-label">Natureza <span class="required-field">*</span></label>
                                    <select class="form-select form-select-sm select2" name="natureza" id="natureza"></select>
                                </div>

                                <div class="col-md-4">
                                    <label for="toleranciaQuantidade" class="form-label">Tolerância de Quantidade +/- <span class="required-field">*</span></label>
                                    <input type="text" class="form-control form-control-sm percentual" name="toleranciaQuantidade" id="toleranciaQuantidade" />
                                </div>

                                <div class="col-md-4">
                                    <label for="toleranciaDias" class="form-label">Tolerância de Dias +/- <span class="required-field">*</span></label>
                                    <input type="text" class="form-control form-control-sm" name="toleranciaDias" id="toleranciaDias" />
                                </div>
                            </div>
                        </div>

                        <div class="tab-content" data-tab-content="2">
                            <div class="section-header">
                                <div class="section-title">Fiscal</div>
                                <p class="section-subtitle">Informações fiscais, tributárias e parâmetros fiscais do Protheus.</p>
                            </div>

                            <div class="row g-3">
                                <div class="col-12 col-md-4">
                                    <label for="contribuinteIcms" class="form-label">Contribuinte de ICMS <span class="required-field">*</span></label>
                                    <select class="form-select form-select-sm select2" name="contribuinteIcms" id="contribuinteIcms">
                                        <option value="">Selecione um item</option>
                                    </select>
                                </div>
                                <div class="col-12 col-md-4">
                                    <label for="icmsDiferido" class="form-label">ICMS Diferido <span class="required-field">*</span></label>
                                    <select class="form-select form-select-sm select2" name="icmsDiferido" id="icmsDiferido">
                                        <option value="0">Não</option>
                                        <option value="1">Sim</option>
                                    </select>
                                </div>
                                <div class="col-12 col-md-4">
                                    <label for="tes" class="form-label">TES</label>
                                    <input disabled type="number" class="form-control form-control-sm" name="tes" id="tes" />
                                </div>
                            </div>
                        </div>

                        <div class="tab-content" data-tab-content="3">
                            <div class="section-header">
                                <div class="section-title">Anexos</div>
                                <p class="section-subtitle">Evidências e documentos relacionados ao cadastro. Para ICMS diferido ou isenção de IPI, anexe os documentos comprobatórios.</p>
                            </div>

                            <div class="row g-3 documentos-pj d-none">
                                <div class="col-md-6 documento-regra documento-pj-obrigatorio d-none">
                                    <document-upload 
                                        name="cartao_cnpj" 
                                        label="Cartão CNPJ" 
                                        file-code="EXTENSAO_CLIENTE_" 
                                        required>
                                    </document-upload>
                                </div>

                                <div class="col-md-6 documento-regra documento-pj-obrigatorio d-none">
                                    <document-upload 
                                        name="comprovante_endereco" 
                                        label="Comprovante de Endereço" 
                                        file-code="EXTENSAO_CLIENTE_" 
                                        required>
                                    </document-upload>
                                </div>

                                <div class="col-md-6 documento-regra d-none">
                                    <document-upload 
                                        name="consulta_sintegra" 
                                        label="Consulta SINTEGRA" 
                                        file-code="EXTENSAO_CLIENTE_" 
                                        required>
                                    </document-upload>
                                </div>

                                <div class="col-md-6 documento-regra documento-pj-obrigatorio d-none">
                                    <document-upload 
                                        name="contrato_social" 
                                        label="Contrato Social e Alterações Contratuais" 
                                        file-code="EXTENSAO_CLIENTE_" 
                                        required>
                                    </document-upload>
                                </div>

                                <div class="col-md-6 documento-regra documento-pj-inscricao-estadual d-none">
                                    <document-upload 
                                        name="inscricao_estadual" 
                                        label="Inscrição Estadual" 
                                        file-code="EXTENSAO_CLIENTE_" 
                                        required>
                                    </document-upload>
                                </div>

                                <div class="col-md-6 documento-regra documento-pj-carta-isencao d-none">
                                    <document-upload 
                                        name="carta_isencao" 
                                        label="Carta de Isenção" 
                                        file-code="EXTENSAO_CLIENTE_" 
                                        required>
                                    </document-upload>
                                </div>

                                <div class="col-md-6 documento-regra documento-pj-deferimento d-none">
                                    <document-upload 
                                        name="documento_deferimento" 
                                        label="Documento Comprobatório de Deferimento" 
                                        file-code="EXTENSAO_CLIENTE_" 
                                        required>
                                    </document-upload>
                                </div>
                            </div>

                            <div class="row g-3 documentos-pf d-none">
                                <div class="col-md-6 documento-regra documento-pf-obrigatorio d-none">
                                    <document-upload 
                                        name="cpf_documento_foto" 
                                        label="CPF e Documento com Foto" 
                                        file-code="EXTENSAO_CLIENTE_" 
                                        required>
                                    </document-upload>
                                </div>

                                <div class="col-md-6 documento-regra documento-pf-obrigatorio d-none">
                                    <document-upload 
                                        name="comprovante_endereco_pf" 
                                        label="Comprovante de Endereço" 
                                        file-code="EXTENSAO_CLIENTE_" 
                                        required>
                                    </document-upload>
                                </div>

                                <div class="col-md-6 documento-regra documento-pf-obrigatorio d-none">
                                    <document-upload 
                                        name="declaracao_ir" 
                                        label="Declaração do Imposto de Renda" 
                                        file-code="EXTENSAO_CLIENTE_" 
                                        required>
                                    </document-upload>
                                </div>

                                <div class="col-md-6 documento-regra documento-pf-inscricao-estadual d-none">
                                    <document-upload 
                                        name="inscricao_estadual_pf" 
                                        label="Inscrição Estadual" 
                                        file-code="EXTENSAO_CLIENTE_" 
                                        required>
                                    </document-upload>
                                </div>

                                <div class="col-md-6 documento-regra documento-pf-carta-isencao d-none">
                                    <document-upload 
                                        name="carta_isencao_pf" 
                                        label="Carta de Isenção" 
                                        file-code="EXTENSAO_CLIENTE_" 
                                        required>
                                    </document-upload>
                                </div>

                                <div class="col-md-6 documento-regra documento-pf-deferimento d-none">
                                    <document-upload 
                                        name="documento_deferimento_pf" 
                                        label="Documento comprobatório de deferimento" 
                                        file-code="EXTENSAO_CLIENTE_" 
                                        required>
                                    </document-upload>
                                </div>
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
</body>
</html>
