<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#Include File="../../cnn2/cnn2.asp"-->
<!--#Include File="../../funcoes/funcoes.asp"-->
<!--#Include File="../../includes/variaveis.asp"-->
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

        .page-actions {
            gap: .5rem;
        }

        .form-container {
            background: #fff;
            border-radius: 8px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }

        .form-label {
            font-weight: 600;
        }

        .required-field {
            color: #dc3545;
        }

        .container {
            max-width: 85%;
        }

        .tabs-container {
            background-color: #f8f9fa;
            border-bottom: 2px solid #e9ecef;
            border-radius: 8px 8px 0 0;
            overflow-x: auto;
        }

        .tabs {
            display: flex;
            padding: 0;
            gap: 0;
            min-width: max-content;
        }

        .tab-button {
            flex: 0 0 auto;
            min-width: 160px;
            padding: 15px 20px;
            background: transparent;
            border: none;
            border-bottom: 3px solid transparent;
            cursor: pointer;
            transition: all 0.3s ease;
            display: flex;
            align-items: center;
            justify-content: center;
            gap: 8px;
            color: #6c757d;
            font-weight: 500;
            white-space: nowrap;
        }

        .tab-button:hover {
            background-color: rgba(0,0,0,0.05);
        }

        .tab-button.active {
            color: #0d6efd;
            border-bottom-color: #0d6efd;
            background-color: #fff;
        }

        .tab-button.completed {
            color: #198754;
        }

        .tab-button.completed .tab-icon {
            color: #198754;
        }

        .tab-button.invalid {
            color: #dc3545;
            border-bottom-color: #dc3545;
        }

        .tab-button.invalid .tab-icon {
            color: #dc3545;
        }

        .tab-button:disabled {
            cursor: not-allowed;
            opacity: 0.5;
        }

        .tab-icon {
            font-size: 16px;
            flex-shrink: 0;
        }

        .tab-title {
            font-size: 14px;
            flex-shrink: 0;
        }

        .tab-content {
            display: none;
            padding: 30px;
            animation: fadeIn 0.3s ease;
            min-height: 300px;
            max-height: 65vh;
            overflow: auto;
        }

        .tab-content.active {
            display: block;
        }

        .z-index-9999 {
            z-index: 9999 !important;
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

        @keyframes fadeIn {
            from {
                opacity: 0;
                transform: translateY(10px);
            }
            to {
                opacity: 1;
                transform: translateY(0);
            }
        }

        .form-navigation {
            padding: 20px 30px;
            background-color: #f8f9fa;
            border-top: 1px solid #e9ecef;
            border-radius: 0 0 8px 8px;
        }

        .loading-icon {
            display: none;
            margin-right: 5px;
        }

        .btn-help {
            padding: 2px 8px;
            font-size: 0.6rem;
            border-radius: 4px;
            display: inline-flex;
            align-items: center;
            gap: 4px;
            text-decoration: none;
            margin-bottom: .37rem;
        }

        .btn-help i {
            font-size: 0.8rem;
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

        @media (max-width: 768px) {
            .tab-button {
                min-width: 140px;
                padding: 12px 15px;
            }

            .tab-title {
                font-size: 13px;
            }
        }

        .anexo-item {
            padding: 12px;
            border: 1px solid #e5e7eb;
            border-radius: 6px;
            margin-bottom: 10px;
            transition: background-color 0.2s;
        }
        .anexo-item:hover {
            background-color: #f9fafb;
        }
        .anexo-icon {
            font-size: 1.5rem;
            color: #3b82f6;
            margin-right: 12px;
        }
        .anexo-info {
            flex: 1;
        }
        .anexo-nome {
            font-weight: 600;
            color: #1f2937;
            margin-bottom: 4px;
        }
        .anexo-data {
            font-size: 0.85rem;
            color: #6b7280;
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

        .form-container {
            overflow: hidden;
        }
        .form-layout {
            display: grid;
            grid-template-columns: 270px minmax(0, 1fr);
        }
        .tabs-container {
            background-color: #f8f9fa;
            border-bottom: none;
            border-radius: 0;
            border-right: 1px solid #e9ecef;
            overflow-y: auto;
            overflow-x: hidden;
        }
        .tabs {
            display: flex;
            flex-direction: column;
            gap: 3px;
            min-width: 0;
            padding: 14px 10px;
        }
        .tab-button {
            align-items: center;
            background: transparent;
            border: none;
            border-radius: 7px;
            color: #495057;
            cursor: pointer;
            display: flex;
            flex: initial;
            font-weight: 500;
            gap: 9px;
            justify-content: flex-start;
            min-width: 0;
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
            border-bottom-color: transparent;
            color: #0d6efd;
            font-weight: 700;
        }
        .tab-button.invalid {
            background-color: #fff5f5;
            border-bottom-color: transparent;
            color: #dc3545;
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
        .tab-content-wrapper {
            background: #fff;
            display: flex;
            flex-direction: column;
            justify-content: space-between;
            min-width: 0;
        }
        .tab-content {
            height: 80vh;
            max-height: none;
            min-height: 0;
            overflow: auto;
        }
        .form-navigation {
            background: #fff;
            border-radius: 0;
            border-top: 1px solid #dee2e6;
            padding: 10px 20px;
            position: sticky;
            bottom: 0;
            z-index: 10;
        }
        
        @media (max-width: 1400px) {
            .form-layout {
                grid-template-columns: 1fr;
            }
            .tabs-container {
                border-bottom: 1px solid #e9ecef;
                border-right: none;
                overflow-x: auto;
                overflow-y: hidden;
            }
            .tabs {
                flex-direction: row;
                min-width: max-content;
                padding: 8px;
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
    <script src="https://unpkg.com/just-validate@latest/dist/just-validate.production.min.js"></script>

    <script src="/corporativo/components/toast.js?v=<%=Now()%>"></script>
    <script src="/corporativo/components/confirmation.js?v=<%=Now()%>"></script>
    <script src="/corporativo/components/upload/fileUpload.js?v=<%=Now()%>"></script>

    <script>
        let abaAtual = 1;
        let totalAbas = 7;
        let abasValidadas = new Set();
        let validacoes = {};
        let modoEdicao = false;
        let modoVisualizacao = false;
        let id = null;
        let codPlanta = ''
        let statusSolicitacao = ''

        let listaAcessorios = [];
        let acessorioEditandoIndex = null;
        
        // Modal de Anexos
        let listaAnexos = [];

        let fichaSelecionada = null;
        let resultadosCopiarFicha = [];

        function exibirLoading(exibir) {
            if (exibir) {
                $('#loadingOverlay').removeClass('d-none').addClass('d-flex');
            } else {
                $('#loadingOverlay').removeClass('d-flex').addClass('d-none');
            }
        }

        function ehTipoCancelamento() {
            return $('#tipoSolicitacao').val() === '3';
        }

        function campoTipoSolicitacaoAlterado() {
            removerAvisoAlteraCancela();
            
            if (ehTipoCancelamento()) {
                $('#textoLabelFTCopia').text('FT a ser Cancelada');
                $('#asteriscoFTCopia').show();
            } else {
                $('#textoLabelFTCopia').text('Copiar de FT');
                $('#asteriscoFTCopia').hide();
            }
        }

        function verificarModoEdicao() {
            const urlParams = new URLSearchParams(window.location.search);
            const paramId = urlParams.get('id');
            const paramCodPlanta = urlParams.get('codPlanta');
            const paramVisualizar = urlParams.get('visualizar');
            
            if (paramId && paramCodPlanta) {
                id = paramId;
                codPlanta = paramCodPlanta;
                $('#idTitulo').text(id);

                if (paramVisualizar === '1') {
                    modoVisualizacao = true;
                } else {
                    modoEdicao = true;
                }
            }
        }

        function aplicarModoVisualizacao() {
            $('#formularioSolicitacaoProjeto input, #formularioSolicitacaoProjeto select, #formularioSolicitacaoProjeto textarea').prop('disabled', true);

            $('.select2').each(function() {
                $(this).prop('disabled', true).trigger('change.select2');
            });

            $('#btnAdicionarAcessorio, #uploadAnexos, #divAcoes').hide();
        }

        function inicializarSelect2() {
            $('.select2').select2({ theme: 'bootstrap-5', width: '100%', dropdownCssClass: "z-index-9999" });
        }

        function inicializarMascaras() {
            $('#amostraCep').mask('00000-000');
            $('#amostraTelefone').mask('(00) 00000-0000');
            $('#amostraCNPJ').mask('00.000.000/0000-00');
        }

        function calcularAreaBlank() {
            const comprimento = $('#comprimentoChapa').val() || 0;
            const largura = $('#larguraChapa').val() || 0;
            $('#areaBlank').val(parseFloat(comprimento) * parseFloat(largura));
        }

        function calcularPeso() {
            const area = $('#areaLiquida').val() || 0;
            const gramatura = $('#gramatura').val() || 0;
            $('#peso').val(parseFloat(area) * parseFloat(gramatura));
        }

        function resetarTodosCampos() {
            $('#codCliente').val('');
            $('#nomeCliente').val('');
            $('#codLoja').val('');
            resetarCamposFormulario();
        }

        function resetarCamposFormulario() {
            removerAvisoAlteraCancela();

            $('#idDesenvolvedor').val('').trigger('change');
            $('#referencia').val('');
            $('#toleranciaMaxima').val('0');
            $('#toleranciaMinima').val('0');
            $('#estilo').val('').trigger('change');
            $('#observacoes').val('').trigger('change');
            $('#comprimento').val('0');
            $('#largura').val('0');
            $('#altura').val('0');
            $('#numeroMatriz').val('0');
            $('#inkJet').val('').trigger('change');
            $('#necessitaAcompanhamento').val('').trigger('change');
            $('#nivelLotePiloto').val('').trigger('change');
            $('#divNivel').hide();
            $('#clienteUsaTara').val('').trigger('change');
            $('#toleranciaPeso').val('0');
            $('#divToleranciaPeso').hide();
            $('#estruturaPapelao').val('').trigger('change');
            $('#gramatura').val('0');
            $('#larguraChapa').val('0');
            $('#comprimentoChapa').val('0');
            $('#areaBlank').val('0');
            $('#areaLiquida').val('0');
            $('#peso').val('0');

            $('#gramaturaReferencial').val('0');
            $('#esmagamento').val('0');
            $('#ectColuna').val('0');
            $('#mullenReferencial').val('0');
            $('#espessura').val('0');
            $('#bctCompressao').val('0');
            $('#cobbInterno').val('0');
            $('#cobbExterno').val('0');

            $('#tipoFechamento').val('').trigger('change');
            $('#pecasPorUnidade').val('').trigger('change');
            $('#ehAmarrado').val('').trigger('change');
            $('#quantidadeFitilhos').val('').trigger('change');
            $('#tipoFitilho').val('').trigger('change');
            $('#divQuantidadeFitilhos').hide();
            $('#divTipoFitilho').hide();
            $('#ehPaletizado').val('').trigger('change');
            $('#camposPaletizado').hide();
            $('#alturaMaximaPalete').val('0');
            $('#palete').val('').trigger('change');
            $('#temFacesOpostas').val('').trigger('change');
            $('#temCantinho').val('').trigger('change');
            $('#temGradeMadeira').val('').trigger('change');
            $('#temFilmeStrech').val('').trigger('change');
            $('#clienteUsaRack').val('').trigger('change');
            $('#overhang').val('').trigger('change');
            $('#quantidadeEspelhos').val('').trigger('change');

            listaAcessorios = [];
            renderizarListaAcessorios();

            $('#conteudoEmbalado').val('').trigger('change');
            $('#consumoMensal').val('0');
            $('#tipoLinhaProducaoCliente').val('').trigger('change');
            $('#tipoProdutoPrimario').val('').trigger('change');
            $('#pesoPrimario').val('0');
            $('#quantidadePrimarios').val('0');
            $('#suportacaoPrimario').val('').trigger('change');
            $('#tempoEstocagem').val('0');
            $('#umidadeRelativa').val('0');
            $('#quantidadeCores').val('0');
            $('#precoKG').val('0');
            $('#orcamento').val('').trigger('change');
            $('#margemContribuicao').val('0');
            $('#resultado').val('0');
            $('#tipoSavings').val('').trigger('change');
            $('#valorSavingSK').val('0');
            $('#valorSavingCliente').val('0');
            $('#valorEstimadoCliche').val('0');
            $('#valorEstimadoMatriz').val('0');
            $('#informacoesAdicionais').val('');

            $('#quantidadeAmostras').val('0').trigger('change');
            $('#camposAmostra').hide();
            $('#amostraCep').val('');
            $('#amostraResponsavel').val('');
            $('#amostraEndereco').val('');
            $('#amostraBairro').val('');
            $('#amostraMunicipio').val('');
            $('#amostraUF').val('').trigger('change');
            $('#amostraCNPJ').val('');
            $('#amostraTelefone').val('');
            $('#amostraEmail').val('');
            $('#amostraObservacoes').val('');

            abasValidadas.clear();
            $('.tab-button').removeClass('completed invalid');
            $('.tab-button .tab-icon').each(function() {
                const tabNum = $(this).closest('.tab-button').data('tab');
                const iconClasses = { 
                    1:'bi bi-box', 
                    2:'bi bi-list-ul', 
                    3:'bi bi-brush', 
                    4:'bi bi-puzzle', 
                    5:'bi bi-archive', 
                    6:'bi bi-clipboard-check' ,
                    7:'bi bi bi-paperclip'
                };
                $(this).attr('class', 'tab-icon ' + (iconClasses[tabNum] || ''));
            });
        }

        function exibirAvisoAlteraCancela(numeroFT, tipoSolicitacao) {
            let texto;

            // Aviso altera/cancela
            if (tipoSolicitacao == 3) {
                texto = `FT <strong>#${numeroFT}</strong> ${(modoEdicao || modoVisualizacao) ? 'foi' : 'será'} cancelada`;
            } else {
                texto = `Copiando dados da FT <strong>#${numeroFT}</strong>`;
            }

            const html = `
                <div class="alert alert-warning alert-dismissible py-2 px-3 mb-0 d-flex align-items-center" role="alert" style="font-size: 0.85rem;">
                    <i class="bi bi-info-circle me-2"></i>
                    <span>${texto}</span>
                </div>
            `;
            $('#avisoAlteraCancela').html(html).show();
            $('#numeroFTCopia').val(numeroFT);
        }

        function removerAvisoAlteraCancela() {
            fichaSelecionada = null;
            $('#avisoAlteraCancela').html('').hide();
            $('#numeroFTCopia').val('');
        }

        function pesquisarCliente() {
            const codPlanta = $('#codPlanta').val();
            const flagPesquisaPreCadastro = $('#flagPesquisaPreCadastro').is(':checked') ? 1 : 0;
            if (!codPlanta) return alert('Selecione a planta primeiro');

            const url = `../cli_busca.asp?Planta=${codPlanta}&paramPesquisaPreCadastro=${flagPesquisaPreCadastro}&paramBuscaNova=1`;
            window.open(url, '', "width=650,height=400");
        }

        function abrirPdfFefco() {
            window.open('https://www.fefco.org/sites/default/files/documents/FEFCO_Codes_of_Designs.pdf', '_blank');
        }

        function abrirModalTiposProduto() {
            const modal = new bootstrap.Modal(document.getElementById('modalTiposProduto'));
            modal.show();
        }

        function abrirModalCopiarFicha() {
            const codCliente = $('#codCliente').val();
            if (!codCliente) { Toast.warning('Informe o código do cliente antes de copiar uma ficha técnica.'); return; }

            fichaSelecionada = null;
            resultadosCopiarFicha = [];
            $('#copiarFichaBusca').val('');
            $('#copiarFichaResultados').html(`
                <div class="text-center text-muted">
                    <i class="bi bi-search" style="font-size: 1.5rem;"></i>
                    <p class="mt-2 mb-0">Digite um termo para buscar fichas técnicas existentes.</p>
                </div>
            `);
            $('#btnConfirmarCopiarFicha').prop('disabled', true);

            const modal = new bootstrap.Modal(document.getElementById('modalCopiarFicha'));
            modal.show();
            setTimeout(function() { $('#copiarFichaBusca').focus(); }, 500);
        }

        function pesquisarFichaTecnica() {
            const busca = $('#copiarFichaBusca').val().trim();
            const codCliente = $('#codCliente').val();
            const codPlanta = $('#codPlanta').val();

            if (!busca) {
                $('#copiarFichaResultados').html(`
                    <div class="text-center text-muted">
                        <i class="bi bi-search" style="font-size: 1.5rem;"></i>
                        <p class="mt-2 mb-0">Digite um termo para buscar fichas técnicas existentes.</p>
                    </div>
                `);
                return;
            }

            fichaSelecionada = null;
            $('#btnConfirmarCopiarFicha').prop('disabled', true);

            $('#copiarFichaResultados').html(`
                <div class="text-center">
                    <div class="spinner-border spinner-border-sm text-primary" role="status">
                        <span class="visually-hidden">Buscando...</span>
                    </div>
                    <p class="mt-2 mb-0 text-muted">Buscando fichas técnicas...</p>
                </div>
            `);

            $.ajax({
                url: 'SolicitacaoProjetoRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: {
                    acao: 'listarFichasTecnicas',
                    filtroCodCliente: codCliente,
                    filtroCodPlanta: codPlanta,
                    busca: busca
                },
                dataType: 'json',
                success: function(response) {
                    if (!response.sucesso) {
                        $('#copiarFichaResultados').html(`
                            <div class="alert text-center alert-danger" role="alert">
                                <i class="bi bi-exclamation-triangle" style="font-size: 1.5rem;"></i>
                                <p class="mt-2 mb-0"> ${response.mensagem || 'Erro ao buscar fichas técnicas.'} </p>
                            </div>
                        `);
                        return;
                    }

                    resultadosCopiarFicha = response.dados || [];

                    if (resultadosCopiarFicha.length === 0) {
                        $('#copiarFichaResultados').html(`
                            <div class="alert text-center alert-info" role="alert">
                                <i class="bi bi-info-circle" style="font-size: 1.5rem;"></i>
                                <p class="mt-2 mb-0"> Nenhuma ficha técnica encontrada com os termos pesquisados. </p>
                            </div>
                        `);
                        return;
                    }

                    renderizarResultadosCopiarFicha();
                },
                error: function(xhr, status, error) {
                    console.error('Erro ao buscar fichas técnicas:', error);
                    $('#copiarFichaResultados').html(`
                        <div class="alert text-center alert-info" role="alert">
                            <i class="bi bi-info-circle" style="font-size: 1.5rem;"></i>
                            <p class="mt-2 mb-0"> Nenhuma ficha técnica encontrada com os termos pesquisados.</p>
                        </div>
                    `);
                }
            });
        }

        function renderizarResultadosCopiarFicha() {
            let html = `
                <div class="table-responsive">
                    <table class="table table-sm table-hover mb-0">
                        <thead class="thead-light">
                            <tr>
                                <th style="width: 50px;"></th>
                                <th>Item</th>
                                <th>Cliente</th>
                                <th>Estilo</th>
                                <th>Referência</th>
                                <th>Estrutura</th>
                            </tr>
                        </thead>
                        <tbody>
            `;

            resultadosCopiarFicha.forEach(function(item, index) {
                html += `
                    <tr data-index="${index}">
                        <td class="text-center">
                            <button onclick="selecionarFichaTecnica(${index})"
                                class="btn-selecionar" 
                                title="Selecionar este item">
                                <i class="bi bi-circle"></i> 
                            </button>
                        </td>
                        <td>${item.item}</td>
                        <td>${item.codCliente}</td>
                        <td>${item.estilo}</td>
                        <td>${item.referencia}</td>
                        <td>${item.estruturaPapelao}</td>
                    </tr>
                `;
            });

            html += `
                        </tbody>
                    </table>
                </div>
            `;

            $('#copiarFichaResultados').html(html);
        }

        function selecionarFichaTecnica(index) {
            fichaSelecionada = resultadosCopiarFicha[index];
            
            $('#copiarFichaResultados .btn-selecionar')
                .removeClass('selecionado')
                .html('<i class="bi bi-circle"></i>')
                .attr('title', 'Selecionar este item');
            
            $('#copiarFichaResultados tr[data-index="' + index + '"] .btn-selecionar')
                .addClass('selecionado')
                .html('<i class="bi bi-check-circle-fill"></i>')
                .attr('title', 'Item Selecionado');
            
            $('#btnConfirmarCopiarFicha').prop('disabled', false);
        }

        function confirmarCopiarFicha() {
            if (!fichaSelecionada) {
                Toast.warning('Selecione uma ficha técnica para copiar.');
                return;
            }

            const modal = bootstrap.Modal.getInstance(document.getElementById('modalCopiarFicha'));
            modal.hide();

            preencherFormularioParaCopia(fichaSelecionada);
        }

        function preencherFormularioParaCopia(dados) {
            exibirAvisoAlteraCancela(dados.item, $('#tipoSolicitacao').val());

            $('#idDesenvolvedor').val(dados.idDesenvolvedor).trigger('change');
            $('#referencia').val(dados.referencia);
            $('#toleranciaMaxima').val(dados.toleranciaMaxima || 0);
            $('#toleranciaMinima').val(dados.toleranciaMinima || 0);
            $('#estilo').val(dados.estilo).trigger('change');
            $('#observacoes').val(dados.observacoes).trigger('change');
            $('#comprimento').val(dados.comprimento || 0);
            $('#largura').val(dados.largura || 0);
            $('#altura').val(dados.altura || 0);
            $('#numeroMatriz').val(dados.numeroMatriz || 0);
            $('#inkJet').val(dados.inkJet).trigger('change');

            $('#necessitaAcompanhamento').val(dados.necessitaAcompanhamento).trigger('change');
            if (dados.necessitaAcompanhamento === 'S') {
                $('#divNivel').show();
                $('#nivelLotePiloto').val(dados.nivelLotePiloto).trigger('change');
            }

            $('#clienteUsaTara').val(dados.clienteUsaTara).trigger('change');
            if (dados.clienteUsaTara == '1') {
                $('#divToleranciaPeso').show();
                $('#toleranciaPeso').val(dados.toleranciaPeso || 0);
            }

            $('#estruturaPapelao').val(dados.estruturaPapelao).trigger('change');
            $('#larguraChapa').val(dados.larguraChapa || 0);
            $('#comprimentoChapa').val(dados.comprimentoChapa || 0);
            $('#areaLiquida').val(dados.areaLiquida || 0);

            calcularAreaBlank();
            calcularPeso();

            $('#gramaturaReferencial').val(dados.gramaturaReferencial || 0);
            $('#esmagamento').val(dados.esmagamento || 0);
            $('#ectColuna').val(dados.ectColuna || 0);
            $('#mullenReferencial').val(dados.mullenReferencial || 0);
            $('#espessura').val(dados.espessura || 0);
            $('#bctCompressao').val(dados.bctCompressao || 0);
            $('#cobbInterno').val(dados.cobbInterno || 0);
            $('#cobbExterno').val(dados.cobbExterno || 0);

            $('#tipoFechamento').val(dados.tipoFechamento).trigger('change');
            $('#pecasPorUnidade').val(dados.pecasPorUnidade).trigger('change');

            $('#ehAmarrado').val(dados.ehAmarrado).trigger('change');
            if (dados.ehAmarrado === 'S') {
                $('#divQuantidadeFitilhos').show();
                $('#divTipoFitilho').show();
                $('#quantidadeFitilhos').val(dados.quantidadeFitilhos).trigger('change');
                $('#tipoFitilho').val(dados.tipoFitilho).trigger('change');
            }

            $('#ehPaletizado').val(dados.ehPaletizado).trigger('change');
            if (dados.ehPaletizado === 'S') {
                $('#camposPaletizado').show();
                $('#alturaMaximaPalete').val(dados.alturaMaximaPalete || 0);
                $('#palete').val(dados.palete).trigger('change');
                $('#temFacesOpostas').val(dados.temFacesOpostas).trigger('change');
                $('#temCantinho').val(dados.temCantinho).trigger('change');
                $('#temGradeMadeira').val(dados.temGradeMadeira).trigger('change');
                $('#temFilmeStrech').val(dados.temFilmeStrech).trigger('change');
                $('#clienteUsaRack').val(dados.clienteUsaRack).trigger('change');
                $('#overhang').val(dados.overhang).trigger('change');
                $('#quantidadeEspelhos').val(dados.quantidadeEspelhos).trigger('change');
            }

            if (dados.acessorios) {
                listaAcessorios = dados.acessorios;
            }

            renderizarListaAcessorios();

            $('#conteudoEmbalado').val(dados.conteudoEmbalado).trigger('change');
            $('#consumoMensal').val(dados.consumoMensal || 0);
            $('#tipoLinhaProducaoCliente').val(dados.tipoLinhaProducaoCliente).trigger('change');
            $('#tipoProdutoPrimario').val(dados.tipoProdutoPrimario).trigger('change');
            $('#pesoPrimario').val(dados.pesoPrimario || 0);
            $('#quantidadePrimarios').val(dados.quantidadePrimarios || 0);
            $('#suportacaoPrimario').val(dados.suportacaoPrimario).trigger('change');
            $('#tempoEstocagem').val(dados.tempoEstocagem || 0);
            $('#umidadeRelativa').val(dados.umidadeRelativa || 0);
            $('#quantidadeCores').val(dados.quantidadeCores || 0);
            $('#precoKG').val(dados.precoKG || 0);
            $('#orcamento').val(dados.orcamento).trigger('change');
            $('#margemContribuicao').val(dados.margemContribuicao || 0);
            $('#resultado').val(dados.resultado || 0);
            $('#tipoSavings').val(dados.tipoSavings).trigger('change');
            $('#valorSavingSK').val(dados.valorSavingSK || 0);
            $('#valorSavingCliente').val(dados.valorSavingCliente || 0);
            $('#valorEstimadoCliche').val(dados.valorEstimadoCliche || 0);
            $('#valorEstimadoMatriz').val(dados.valorEstimadoMatriz || 0);
            $('#informacoesAdicionais').val(dados.informacoesAdicionais);

            $('#quantidadeAmostras').val(dados.quantidadeAmostras || 0).trigger('change');
            var qtdAmostras = parseInt(dados.quantidadeAmostras) || 0;

            if (qtdAmostras >= 1) {
                $('#camposAmostra').show();
                $('#amostraCep').val(dados.amostraCep).trigger('input');
                $('#amostraResponsavel').val(dados.amostraResponsavel);
                $('#amostraEndereco').val(dados.amostraEndereco);
                $('#amostraBairro').val(dados.amostraBairro);
                $('#amostraMunicipio').val(dados.amostraMunicipio);
                $('#amostraUF').val(dados.amostraUF).trigger('change');
                $('#amostraCNPJ').val(dados.amostraCNPJ).trigger('input');
                $('#amostraTelefone').val(dados.amostraTelefone).trigger('input');
                $('#amostraEmail').val(dados.amostraEmail);
                $('#amostraObservacoes').val(dados.amostraObservacoes);
            }

            Toast.success('Dados da ficha técnica copiados com sucesso!');
        }

        function carregarDesenvolvedores() {
            const $select = $('#idDesenvolvedor');
            $select.empty().append('<option value="">Selecione o desenvolvedor...</option>');

            const filtroCodPlanta = $('#codPlanta').val();
            if (!filtroCodPlanta) return $.Deferred().resolve().promise();

            const params = {
                acao: 'listarDesenvolvedores',
                filtroCodPlanta: filtroCodPlanta
            }

            return $.ajax({
                url: 'SolicitacaoProjetoRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: params,
                dataType: 'json',
                success: function(response) {
                    if (!response.sucesso) {
                        console.error("Erro ao carregar desenvolvedores: " + response.mensagem)
                        Toast.error("Erro ao carregar desenvolvedores: " + response.mensagem)
                        return;
                    }
                    response.dados.forEach(function(item) {
                        $select.append($('<option>', { value: item.id, text: item.nome }));
                    });
                },
                error: function(xhr, status, error) {
                    console.error('Erro ao carregar desenvolvedores:', error);
                    Toast.error('Erro ao carregar desenvolvedores');
                }
            });
        }

        function carregarEstilos() {
            const $select = $('#estilo');
            $select.empty().append('<option value="">Selecione o estilo...</option>');

            const $selectModal = $('#acessorioEstilo');
            $selectModal.empty().append('<option value="">Selecione o estilo...</option>');

            const filtroCodPlanta = $('#codPlanta').val();
            if (!filtroCodPlanta) return $.Deferred().resolve().promise();

            const params = {
                acao: 'listarEstilos',
                filtroCodPlanta: filtroCodPlanta
            }

            return $.ajax({
                url: 'SolicitacaoProjetoRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: params,
                dataType: 'json',
                success: function(response) {
                    if (!response.sucesso) {
                        console.error("Erro ao carregar estilos: " + response.mensagem)
                        Toast.error("Erro ao carregar estilos: " + response.mensagem)
                        return;
                    }
                    response.dados.forEach(function(item) {
                        const option = { value: item.codigo, text: `${item.codigo} - ${item.descricao}` }
                        $select.append($('<option>', option));
                        $selectModal.append($('<option>', option));
                    });
                },
                error: function(xhr, status, error) {
                    console.error('Erro ao carregar estilos:', error);
                    Toast.error('Erro ao carregar estilos');
                }
            });
        }

        function carregarEstruturasPapelao() {
            const $select = $('#estruturaPapelao');
            $select.empty().append('<option value="">Selecione a estrutura...</option>');
            
            const $selectModal = $('#acessorioEstrutura');
            $selectModal.empty().append('<option value="">Selecione a estrutura...</option>');

            const filtroCodPlanta = $('#codPlanta').val();
            if (!filtroCodPlanta) return $.Deferred().resolve().promise();

            const params = {
                acao: 'listarEstruturasPapelao',
                filtroCodPlanta: filtroCodPlanta
            }

            return $.ajax({
                url: 'SolicitacaoProjetoRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: params,
                dataType: 'json',
                success: function(response) {
                    if (!response.sucesso) {
                        console.error('Erro ao carregar estruturas de papelão:', response.mensagem);
                        Toast.error('Erro ao carregar estruturas de papelão');
                        return;
                    }
                    response.dados.forEach(function(item) {
                        const option = { value: item.codigo, text: item.codigo, 'data-gramatura': item.gramatura }
                        $select.append($('<option>', option));
                        $selectModal.append($('<option>', option));
                    });
                },
                error: function(xhr, status, error) {
                    console.error('Erro ao carregar estruturas de papelão:', error);
                    Toast.error('Erro ao carregar estruturas de papelão');
                }
            });
        }

        function carregarConteudosEmbalado() {
            const $select = $('#conteudoEmbalado');
            $select.empty().append('<option value="">Selecione o conteúdo...</option>');

            const filtroCodPlanta = $('#codPlanta').val();
            if (!filtroCodPlanta) return $.Deferred().resolve().promise();

            const params = {
                acao: 'listarConteudosEmbalado',
                filtroCodPlanta: filtroCodPlanta
            }

            return $.ajax({
                url: 'SolicitacaoProjetoRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: params,
                dataType: 'json',
                success: function(response) {
                    if (!response.sucesso) {
                        console.error("Erro ao carregar conteúdos: " + response.mensagem)
                        Toast.error("Erro ao carregar conteúdos: " + response.mensagem)
                        return;
                    }
                    response.dados.forEach(function(item) {
                        $select.append($('<option>', { value: item.codigo, text: item.descricao }));
                    });
                },
                error: function(xhr, status, error) {
                    console.error('Erro ao carregar conteúdos:', error);
                    Toast.error('Erro ao carregar conteúdos');
                }
            });
        }

        function carregarTiposProdutoPrimario() {
            const $select = $('#tipoProdutoPrimario');
            $select.empty().append('<option value="">Selecione o tipo...</option>');

            const filtroCodPlanta = $('#codPlanta').val();
            if (!filtroCodPlanta) return $.Deferred().resolve().promise();

            const params = {
                acao: 'listarTiposProdutoPrimario',
                filtroCodPlanta: filtroCodPlanta
            }

            return $.ajax({
                url: 'SolicitacaoProjetoRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: params,
                dataType: 'json',
                success: function(response) {
                    if (!response.sucesso) {
                        console.error("Erro ao carregar tipos de produto primário: " + response.mensagem)
                        Toast.error("Erro ao carregar tipos de produto primário: " + response.mensagem)
                        return;
                    }
                    response.dados.forEach(function(item) {
                        $select.append($('<option>', { value: item.codigo, text: item.descricao }));
                    });
                },
                error: function(xhr, status, error) {
                    console.error('Erro ao carregar tipos de produto primário:', error);
                    Toast.error('Erro ao carregar tipos de produto primário');
                }
            });
        }

        function carregarPaletes() {
            const $select = $('#palete');
            $select.empty().append('<option value="">Selecione o palete...</option>');

            const filtroCodPlanta = $('#codPlanta').val();
            if (!filtroCodPlanta) return $.Deferred().resolve().promise();

            const params = {
                acao: 'listarPaletes',
                filtroCodPlanta: filtroCodPlanta
            }

            return $.ajax({
                url: 'SolicitacaoProjetoRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: params,
                dataType: 'json',
                success: function(response) {
                    if (!response.sucesso) {
                        console.error("Erro ao carregar paletes: " + response.mensagem)
                        Toast.error("Erro ao carregar paletes: " + response.mensagem)
                        return;
                    }
                    response.dados.forEach(function(item) {
                        $select.append($('<option>', { value: item.chave, text: item.descricao }));
                    });
                },
                error: function(xhr, status, error) {
                    console.error('Erro ao carregar paletes:', error);
                    Toast.error('Erro ao carregar paletes');
                }
            });
        }

        function campoPlantaAlterado() {
            if (!modoEdicao && !modoVisualizacao) { 
                resetarTodosCampos(); 
            }
            carregarDesenvolvedores(); 
            carregarEstilos(); 
            carregarEstruturasPapelao(); 
            carregarConteudosEmbalado(); 
            carregarTiposProdutoPrimario(); 
            carregarPaletes();
        }

        function campoClienteAlterado() {
            if (!modoEdicao && !modoVisualizacao) { 
                $('#numeroFTCopia').val('');
                resetarCamposFormulario(); 
            }
        }

        function campoNecessitaAcompanhamentoAlterado() {
            const necessitaAcompanhamento = $('#necessitaAcompanhamento').val();
            const $divNivel = $('#divNivel');
            const $nivelLotePiloto = $('#nivelLotePiloto');

            if (necessitaAcompanhamento === 'S') {
                $divNivel.show();
            } else {
                $divNivel.hide();
                $nivelLotePiloto.val('').trigger('change');
            }
        }

        function campoEstruturaPapelaoAlterado() {
            const gramatura = $(this).find('option:selected').attr('data-gramatura') || 0;
            $('#gramatura').val(gramatura);
            $('#gramatura').trigger('change');
        }

        function campoAmarradoAlterado() {
            const ehAmarrado = $('#ehAmarrado').val();
            const $divQuantidadeFitilhos = $('#divQuantidadeFitilhos');
            const $divTipoFitilho = $('#divTipoFitilho');

            if (ehAmarrado === 'S') {
                $divQuantidadeFitilhos.show();
                $divTipoFitilho.show();
            } else {
                $divQuantidadeFitilhos.hide();
                $divTipoFitilho.hide();
                $('#quantidadeFitilhos').val('').trigger('change');
                $('#tipoFitilho').val('').trigger('change');
            }
        }

        function campoClienteUsaTaraAlterado() {
            const clienteUsaTara = $('#clienteUsaTara').val();
            const $divToleranciaPeso = $('#divToleranciaPeso');

            if (clienteUsaTara == '1') {
                $divToleranciaPeso.show();
            } else {
                $divToleranciaPeso.hide();
                $('#toleranciaPeso').val('');
            }
        }

        function campoQuantidadeAmostrasAlterado() {
            const quantidade = parseInt($('#quantidadeAmostras').val()) || 0;
            const clienteEndereco = ($('#hidEndereco').val() || '');
            const clienteBairro = ($('#hidBairro').val() || '');
            const clienteEmail = ($('#hidEmail').val() || '');
            const clienteMunicipio = ($('#hidMunicipio').val() || '');
            const clienteUF = ($('#hidUF').val() || '');
            const clienteCEP = ($('#hidCEP').val() || '')?.replace(/\D/g, '');
            const clienteTelefone = ($('#hidTelefone').val() || '')?.replace(/\D/g, '');
            const clienteCNPJ = ($('#hidCNPJ').val() || '')?.replace(/\D/g, '');
            
            const $camposAmostra = $('#camposAmostra');

            if (quantidade >= 1) {
                $camposAmostra.slideDown(300);
                $('#amostraCep').val(clienteCEP).trigger('input');
                $('#amostraEndereco').val(clienteEndereco).trigger('input');
                $('#amostraBairro').val(clienteBairro).trigger('input');
                $('#amostraMunicipio').val(clienteMunicipio).trigger('input');
                $('#amostraUF').val(clienteUF).trigger('change');
                $('#amostraCNPJ').val(clienteCNPJ).trigger('input');
                $('#amostraTelefone').val(clienteTelefone).trigger('input');
                $('#amostraEmail').val(clienteEmail).trigger('input');
            } else {
                $camposAmostra.slideUp(300);
                $('#amostraCep').val('');
                $('#amostraResponsavel').val('');
                $('#amostraEndereco').val('');
                $('#amostraBairro').val('');
                $('#amostraMunicipio').val('');
                $('#amostraUF').val('').trigger('change');
                $('#amostraCNPJ').val('');
                $('#amostraTelefone').val('');
                $('#amostraEmail').val('');
                $('#amostraObservacoes').val('');
            }
        }

        function campoPaletizadoAlterado() {
            const ehPaletizado = $('#ehPaletizado').val();
            const $camposPaletizado = $('#camposPaletizado');

            if (ehPaletizado === 'S') {
                $camposPaletizado.slideDown(300);
            } else {
                $('#alturaMaximaPalete').val('');
                $('#palete').val('').trigger('change');
                $('#temFacesOpostas').val('').trigger('change');
                $('#temCantinho').val('').trigger('change');
                $('#temGradeMadeira').val('').trigger('change');
                $('#temFilmeStrech').val('').trigger('change');
                $('#clienteUsaRack').val('').trigger('change');
                $('#overhang').val('').trigger('change');
                $('#quantidadeEspelhos').val('').trigger('change');
                $camposPaletizado.slideUp(300);
            }
        }
       
        function buscarCep() {
            const cep = $('#amostraCep').val().replace(/\D/g, '');
            const $loading = $('#cepLoading');
            
            if (cep.length !== 8) {
                return;
            }

            $loading.show();
            
            $('#amostraMunicipio').val('');
            $('#amostraUF').val('').trigger('change');

            $.ajax({
                url: `https://viacep.com.br/ws/${cep}/json/`,
                type: 'GET',
                dataType: 'json',
                timeout: 10000,
                success: function(response) {
                    $loading.hide();
                    
                    if (response.erro) {
                        Toast.warning('CEP não encontrado. Preencha os campos manualmente.');
                        $('#amostraEndereco').focus();
                        return;
                    }

                    $('#amostraMunicipio').val(response.localidade || '');
                    $('#amostraUF').val(response.uf || '').trigger('change');
                    $('#amostraResponsavel').focus();
                },
                error: function(xhr, status, error) {
                    $loading.hide();
                    console.error('Erro ao buscar CEP:', error);
                    Toast.error('Erro ao buscar CEP. Preencha os campos manualmente.');
                    $('#amostraEndereco').focus();
                }
            });
        }

        function abrirModalAcessorio() {
            acessorioEditandoIndex = null;
            $('#formularioAcessorio')[0].reset();
            $('#modalAcessorioTitulo').text('Adicionar Acessório');
            $('#btnSalvarAcessorio').html('<i class="bi bi-plus-lg me-1"></i>Adicionar');
            
            const modal = new bootstrap.Modal(document.getElementById('modalAcessorio'));
            modal.show();
        }

        function editarAcessorio(index) {
            const acessorio = listaAcessorios[index];
            if (!acessorio) return;

            acessorioEditandoIndex = index;
            
            $('#acessorioQuantidade').val(acessorio.quantidade);
            $('#acessorioNumeroCortes').val(acessorio.numeroCortes);
            $('#acessorioLargura').val(acessorio.largura);
            $('#acessorioComprimento').val(acessorio.comprimento);
            $('#acessorioEstrutura').val(acessorio.estrutura).trigger('change');
            $('#acessorioEstilo').val(acessorio.estilo).trigger('change');

            $('#modalAcessorioTitulo').text('Editar Acessório');
            $('#btnSalvarAcessorio').html('<i class="bi bi-check-lg me-1"></i>Salvar');

            const modal = new bootstrap.Modal(document.getElementById('modalAcessorio'));
            modal.show();
        }

        function removerAcessorio(index) {
            ConfirmationModal.danger({
                message: 'Deseja realmente remover este acessório?',
                type: 'danger',
                confirmText: 'Remover',
                cancelText: 'Cancelar',
                onConfirm: () => {
                    listaAcessorios.splice(index, 1);
                    renderizarListaAcessorios();
                    Toast.success('Acessório removido com sucesso!');
                },
                onCancel: () => {}
            });
        }

        function salvarAcessorio() {
            const $quantidade = $('#acessorioQuantidade');
            const $estilo = $('#acessorioEstilo');
            const $numeroCortes = $('#acessorioNumeroCortes');
            const $largura = $('#acessorioLargura');
            const $comprimento = $('#acessorioComprimento');
            const $estrutura = $('#acessorioEstrutura');

            const acessorio = {
                quantidade: $quantidade.val(),
                estilo: $estilo.val(),
                numeroCortes: parseInt($numeroCortes.val()).toString(),
                largura: parseFloat($largura.val()).toString(),
                comprimento: parseFloat($comprimento.val()).toString(),
                estrutura: $estrutura.val()
            };

            if (acessorioEditandoIndex !== null) {
                listaAcessorios[acessorioEditandoIndex] = acessorio;
                Toast.success('Acessório atualizado com sucesso!');
            } else {
                listaAcessorios.push(acessorio);
                Toast.success('Acessório adicionado com sucesso!');
            }

            $quantidade.val('');
            $numeroCortes.val('');
            $largura.val('');
            $comprimento.val('');
            $estrutura.val('').trigger('change');
            $estilo.val('').trigger('change');
            renderizarListaAcessorios();
            
            const modal = bootstrap.Modal.getInstance(document.getElementById('modalAcessorio'));
            modal.hide();
        }

        function renderizarListaAcessorios() {
            const $container = $('#acessoriosListaContainer');
            const $contador = $('#contadorAcessorios');
            
            $contador.text(listaAcessorios.length);

            if (listaAcessorios.length === 0) {
                $container.html(`
                    <div class="text-center py-5">
                        <p class="text-muted mt-3 mb-0">Nenhum acessório adicionado.</p>
                    </div>
                `);
                return;
            }

            let html = `
                <div class="table-responsive">
                    <table class="table table-hover table-sm mb-0">
                        <thead class="table-light">
                            <tr>
                                <th class="text-center" style="width: 50px;">#</th>
                                <th class="text-center">Qtd</th>
                                <th>Estilo</th>
                                <th class="text-center">Nº Cortes</th>
                                <th class="text-center">Largura (mm)</th>
                                <th class="text-center">Comprimento (mm)</th>
                                <th class="text-center">Estrutura</th>
                                ${!modoVisualizacao ? '<th class="text-center" style="width: 100px;">Ações</th>' : ''}
                            </tr>
                        </thead>
                        <tbody>
            `;

            listaAcessorios.forEach((acessorio, index) => {
                html += `
                    <tr>
                        <td class="text-center">${index + 1}</td>
                        <td class="text-center">
                            <span class="badge bg-secondary">${acessorio.quantidade}</span>
                        </td>
                        <td>${acessorio.estiloDescricao || acessorio.estilo}</td>
                        <td class="text-center">${acessorio.numeroCortes}</td>
                        <td class="text-center">${acessorio.largura}</td>
                        <td class="text-center">${acessorio.comprimento}</td>
                        <td class="text-center"><code>${acessorio.estrutura}</code></td>
                        ${!modoVisualizacao ? `
                        <td class="text-center">
                            <button type="button" class="btn btn-sm btn-primary me-1" onclick="editarAcessorio(${index})" title="Editar">
                                <i class="bi bi-pencil"></i>
                            </button>
                            <button type="button" class="btn btn-sm btn-danger" onclick="removerAcessorio(${index})" title="Remover">
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

        /* ANEXOS */
        function renderizarListaAnexos() {
            const $container = $('#anexosListaContainer');
            
            if (listaAnexos.length === 0) {
                $container.html(``);
                return;
            }


            let html = `
                <div class="anexos-list mt-4">    
            `;

            listaAnexos.forEach((anexo, index) => {
                html += `
                    <div class="anexo-item d-flex align-items-center">
                        <i class="bi bi-file-earmark anexo-icon"></i>
                        <div class="anexo-info">
                            <div class="anexo-nome">${anexo.nome}</div>
                            <div class="anexo-data">Adicionado em: ${anexo.dataCriacao}</div>
                        </div>
                        <a href="/Clientes_Docs/${anexo.caminho}" target="_blank" class="btn btn-sm btn-primary">
                            <i class="bi bi-box-arrow-up-right me-1"></i> Abrir
                        </a>
                    </div>
                `;
            });

            html += `
                </div>
            `;

            $container.html(html);
        }

        function removerAnexo(index) {
            ConfirmationModal.danger({
                message: 'Deseja realmente remover este Anexo?',
                type: 'danger',
                confirmText: 'Remover',
                cancelText: 'Cancelar',
                onConfirm: () => {
                    listaAnexos.splice(index, 1);
                    renderizarListaAnexos();
                    Toast.success('Anexo removido com sucesso!');
                },
                onCancel: () => {}
            });
        }

        function inicializarValidacoes() {
            const naoVazio = (msg) => ({
                validator: (value) => value !== null && value !== undefined && String(value).trim() !== '',
                errorMessage: msg
            });

            const condRequerido = (condicao, msg) => ({
                validator: (value) => !condicao() || (value !== null && value !== undefined && String(value).trim() !== ''),
                errorMessage: msg
            });

            validacoes.aba1 = new JustValidate('#formularioSolicitacaoProjeto', {
                lockForm: false, 
            });
            validacoes.aba1
                .addField('#codPlanta', [
                    { rule: 'required', errorMessage: 'Planta é obrigatória' }
                ])
                .addField('#tipoSolicitacao', [
                    { rule: 'required', errorMessage: 'Tipo de solicitação é obrigatório' }
                ])
                .addField('#numeroFTCopia', [
                    condRequerido(
                        () => ehTipoCancelamento(),
                        'FT a ser cancelada é obrigatória'
                    )
                ])
                .addField('#idDesenvolvedor', [
                    { rule: 'required', errorMessage: 'Técnico/Desenvolvedor é obrigatório' }
                ])
                .addField('#referencia', [
                    { rule: 'required', errorMessage: 'Referência é obrigatória' }
                ])
                .addField('#necessitaAcompanhamento', [
                    { rule: 'required', errorMessage: 'Campo obrigatório' }
                ])
                .addField('#nivelLotePiloto', [
                    condRequerido(
                        () => $('#necessitaAcompanhamento').val() === 'S',
                        'Nível é obrigatório quando necessita acompanhamento'
                    )
                ])
                .addField('#toleranciaMaxima', [naoVazio('Tolerância máxima é obrigatória')])
                .addField('#toleranciaMinima', [naoVazio('Tolerância mínima é obrigatória')])
                .addField('#estilo', [
                    { rule: 'required', errorMessage: 'Estilo é obrigatório' }
                ])
                .addField('#comprimento',      [naoVazio('Comprimento é obrigatório')])
                .addField('#largura',          [naoVazio('Largura é obrigatória')])
                .addField('#altura',           [naoVazio('Altura é obrigatória')])
                .addField('#numeroMatriz',     [naoVazio('Número da matriz é obrigatório')])
                .addField('#clienteUsaTara', [
                    { rule: 'required', errorMessage: 'Campo obrigatório' }
                ])
                .addField('#toleranciaPeso', [
                    condRequerido(
                        () => $('#clienteUsaTara').val() === '1',
                        'Tolerância de peso é obrigatória'
                    )
                ])
                .addField('#estruturaPapelao', [
                    { rule: 'required', errorMessage: 'Estrutura de papelão é obrigatória' }
                ])
                .addField('#larguraChapa',     [naoVazio('Largura da chapa é obrigatória')])
                .addField('#comprimentoChapa', [naoVazio('Comprimento da chapa é obrigatório')])
                .addField('#areaLiquida',      [naoVazio('Área líquida é obrigatória')])
                .addField('#inkJet', [
                    { rule: 'required', errorMessage: 'InkJet é obrigatório' }
                ]);

            validacoes.aba2 = new JustValidate('#formularioSolicitacaoProjeto', {
                lockForm: false, 
            });
            validacoes.aba2
                .addField('#gramaturaReferencial', [naoVazio('Gramatura referencial é obrigatória')])
                .addField('#esmagamento',          [naoVazio('Esmagamento é obrigatório')])
                .addField('#ectColuna',            [naoVazio('ECT/Coluna é obrigatório')])
                .addField('#mullenReferencial',    [naoVazio('Mullen referencial é obrigatório')])
                .addField('#espessura',            [naoVazio('Espessura é obrigatória')])
                .addField('#bctCompressao',        [naoVazio('BCT/Compressão é obrigatório')])
                .addField('#cobbInterno',          [naoVazio('COBB Interno é obrigatório')])
                .addField('#cobbExterno',          [naoVazio('COBB Externo é obrigatório')]);

            validacoes.aba3 = new JustValidate('#formularioSolicitacaoProjeto', {
                lockForm: false, 
            });
            validacoes.aba3
                .addField('#tipoFechamento', [
                    { rule: 'required', errorMessage: 'Fechamento é obrigatório' }
                ])
                .addField('#ehAmarrado', [
                    { rule: 'required', errorMessage: 'Amarrado é obrigatório' }
                ])
                .addField('#quantidadeFitilhos', [
                    condRequerido(
                        () => $('#ehAmarrado').val() === 'S',
                        'Quantidade de fitilhos é obrigatória'
                    )
                ])
                .addField('#tipoFitilho', [
                    condRequerido(
                        () => $('#ehAmarrado').val() === 'S',
                        'Tipo de fitilho é obrigatório'
                    )
                ])
                .addField('#pecasPorUnidade', [
                    { rule: 'required', errorMessage: 'Pacotes de é obrigatório' }
                ])
                .addField('#ehPaletizado', [
                    { rule: 'required', errorMessage: 'Paletizado é obrigatório' }
                ])
                .addField('#palete', [
                    condRequerido(
                        () => $('#ehPaletizado').val() === 'S',
                        'Palete é obrigatório'
                    )
                ])
                .addField('#alturaMaximaPalete', [
                    condRequerido(
                        () => $('#ehPaletizado').val() === 'S',
                        'Altura máxima do palete é obrigatória'
                    )
                ])
                .addField('#temFacesOpostas', [
                    condRequerido(
                        () => $('#ehPaletizado').val() === 'S',
                        'Faces opostas é obrigatório'
                    )
                ])
                .addField('#temCantinho', [
                    condRequerido(
                        () => $('#ehPaletizado').val() === 'S',
                        'Cantinho é obrigatório'
                    )
                ])
                .addField('#temGradeMadeira', [
                    condRequerido(
                        () => $('#ehPaletizado').val() === 'S',
                        'Grade de madeira é obrigatório'
                    )
                ])
                .addField('#temFilmeStrech', [
                    condRequerido(
                        () => $('#ehPaletizado').val() === 'S',
                        'Filme strech é obrigatório'
                    )
                ])
                .addField('#clienteUsaRack', [
                    condRequerido(
                        () => $('#ehPaletizado').val() === 'S',
                        'Cliente usa rack é obrigatório'
                    )
                ])
                .addField('#overhang', [
                    condRequerido(
                        () => $('#ehPaletizado').val() === 'S',
                        'Overhang é obrigatório'
                    )
                ])
                .addField('#quantidadeEspelhos', [
                    condRequerido(
                        () => $('#ehPaletizado').val() === 'S',
                        'Espelho é obrigatório'
                    )
                ]);

            validacoes.aba5 = new JustValidate('#formularioSolicitacaoProjeto', {
                lockForm: false, 
            });
            validacoes.aba5
                .addField('#conteudoEmbalado', [
                    { rule: 'required', errorMessage: 'Conteúdo a ser embalado é obrigatório' }
                ])
                .addField('#consumoMensal',           [naoVazio('Consumo mensal é obrigatório')])
                .addField('#tipoLinhaProducaoCliente', [
                    { rule: 'required', errorMessage: 'Linha de produção do cliente é obrigatória' }
                ])
                .addField('#tipoProdutoPrimario', [
                    { rule: 'required', errorMessage: 'Tipo de produto primário é obrigatório' }
                ])
                .addField('#pesoPrimario',        [naoVazio('Peso do primário é obrigatório')])
                .addField('#quantidadePrimarios', [naoVazio('Quantidade de primários é obrigatória')])
                .addField('#suportacaoPrimario', [
                    { rule: 'required', errorMessage: 'Suportação do primário é obrigatória' }
                ])
                .addField('#tempoEstocagem',  [naoVazio('Tempo de estocagem é obrigatório')])
                .addField('#umidadeRelativa', [naoVazio('Umidade relativa é obrigatória')])
                .addField('#quantidadeCores', [naoVazio('Quantidade de cores é obrigatória')])
                .addField('#precoKG',         [naoVazio('Preço KG é obrigatório')])
                .addField('#margemContribuicao',  [naoVazio('Margem de contribuição é obrigatória')])
                .addField('#resultado',           [naoVazio('Resultado é obrigatório')])
                .addField('#tipoSavings', [
                    { rule: 'required', errorMessage: 'Savings é obrigatório' }
                ])
                .addField('#valorSavingSK',       [naoVazio('Valor Saving SK é obrigatório')])
                .addField('#valorSavingCliente',  [naoVazio('Valor Saving Cliente é obrigatório')])
                .addField('#valorEstimadoCliche', [naoVazio('Valor estimado clichê é obrigatório')])
                .addField('#valorEstimadoMatriz', [naoVazio('Valor estimado matriz é obrigatório')]);

            validacoes.aba6 = new JustValidate('#formularioSolicitacaoProjeto', {
                lockForm: false, 
            });
            validacoes.aba6
                .addField('#quantidadeAmostras', [
                    { rule: 'required', errorMessage: 'Quantidade de amostras é obrigatória' }
                ])
                .addField('#amostraCep', [
                    condRequerido(
                        () => parseInt($('#quantidadeAmostras').val()) >= 1,
                        'CEP é obrigatório'
                    )
                ])
                .addField('#amostraResponsavel', [
                    condRequerido(
                        () => parseInt($('#quantidadeAmostras').val()) >= 1,
                        'Responsável é obrigatório'
                    )
                ])
                .addField('#amostraTelefone', [
                    condRequerido(
                        () => parseInt($('#quantidadeAmostras').val()) >= 1,
                        'Telefone é obrigatório'
                    )
                ])
                .addField('#amostraEndereco', [
                    condRequerido(
                        () => parseInt($('#quantidadeAmostras').val()) >= 1,
                        'Endereço de entrega é obrigatório'
                    )
                ])
                .addField('#amostraBairro', [
                    condRequerido(
                        () => parseInt($('#quantidadeAmostras').val()) >= 1,
                        'Bairro é obrigatório'
                    )
                ])
                .addField('#amostraMunicipio', [
                    condRequerido(
                        () => parseInt($('#quantidadeAmostras').val()) >= 1,
                        'Município é obrigatório'
                    )
                ])
                .addField('#amostraUF', [
                    condRequerido(
                        () => parseInt($('#quantidadeAmostras').val()) >= 1,
                        'UF é obrigatória'
                    )
                ])
                .addField('#amostraCNPJ', [
                    condRequerido(
                        () => parseInt($('#quantidadeAmostras').val()) >= 1,
                        'CNPJ é obrigatório'
                    )
                ])
                .addField('#amostraEmail', [
                    condRequerido(
                        () => parseInt($('#quantidadeAmostras').val()) >= 1,
                        'E-mail é obrigatório'
                    )
                ]);

            // Aba Anexos, sem campos obrigatórios
            validacoes.aba7 = new JustValidate('#formularioSolicitacaoProjeto', {
                lockForm: false, 
            });

            validacoes.modalAcessorio = new JustValidate('#formularioAcessorio');
            validacoes.modalAcessorio
                .addField('#acessorioQuantidade', [
                    { rule: 'required', errorMessage: 'Quantidade é obrigatória' }
                ])
                .addField('#acessorioEstilo', [
                    { rule: 'required', errorMessage: 'Estilo é obrigatório' }
                ])
                .addField('#acessorioNumeroCortes', [
                    { rule: 'required', errorMessage: 'Número de cortes é obrigatório' }
                ])
                .addField('#acessorioLargura', [
                    { rule: 'required', errorMessage: 'Largura é obrigatória' }
                ])
                .addField('#acessorioComprimento', [
                    { rule: 'required', errorMessage: 'Comprimento é obrigatório' }
                ])
                .addField('#acessorioEstrutura', [
                    { rule: 'required', errorMessage: 'Estrutura é obrigatória' }
                ])
                .onFail((fields) => {
                    console.log('Formulário de acessório inválido', fields);
                });
        }

        function navegarParaGrade() {
            window.location.href = `SolicitacaoProjetoGrade.asp?codPlanta=${codPlanta}`;
        }

         function navegarParaAba(numeroAba) {
            if (numeroAba < 1 || numeroAba > totalAbas) return;

            abaAtual = numeroAba;

            $('.tab-button').removeClass('active');
            $(`.tab-button[data-tab="${numeroAba}"]`).addClass('active');

            $('.tab-content').removeClass('active');
            $(`.tab-content[data-tab-content="${numeroAba}"]`).addClass('active');
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

        function validarAbaAtual() {
            return new Promise((resolve) => {
                const validator = validacoes[`aba${abaAtual}`];
                
                if (!validator) {
                    resolve(true);
                    return;
                }

                validator.revalidate().then(isValid => {
                    if (!isValid) {
                        const errorsParaExibir = {};
                        Object.entries(validator.fields).forEach(([index, campo]) => {
                            if(campo.isValid) return;

                            const primeiroErro = campo.rules.find(rule => !rule.isValid);
                            if (primeiroErro) {
                                errorsParaExibir[`#${campo.elem.id}`] = primeiroErro.errorMessage || 'Campo inválido';
                            }   
                        });

                        validator.showErrors(errorsParaExibir);
                    }
                    resolve(isValid);
                });
            });
        }

        function exibirErrosValidacao(validator) {
            if (!validator || !validator.fields) return;

            const errorsParaExibir = {};
            Object.entries(validator.fields).forEach(([index, campo]) => {
                if (campo.isValid) return;

                const primeiroErro = campo.rules.find(rule => !rule.isValid);
                if (primeiroErro && campo.elem && campo.elem.id) {
                    errorsParaExibir[`#${campo.elem.id}`] = primeiroErro.errorMessage || 'Campo inválido';
                }
            });

            validator.showErrors(errorsParaExibir);
        }

        function carregarDadosEdicao() {
            const params = { 
                acao: 'obterPorId', 
                id: id, 
                codPlanta: codPlanta
            }

            $.ajax({
                url: 'SolicitacaoProjetoRepositorio.asp',
                type: 'GET',
                data: params,
                dataType: 'json',
                success: function(response) {
                    if (!response.sucesso) {
                        Toast.error(response.mensagem);
                        return;
                    }

                    statusSolicitacao = response.dados.status;
                    preencherFormulario(response.dados);
                },
                error: function(xhr, status, error) {
                    console.error('Erro ao carregar dados para edição:', error);
                    Toast.error('Erro ao carregar dados para edição');
                }
            });
        }

        function preencherFormulario(dados) {
            $('#codPlanta').prop('disabled', true)
            $('#codCliente').prop('disabled', true)
            $('#btnBuscarCliente').prop('disabled', true)
            $('#btnCopiarFicha').prop('disabled', true)
            $('#flagPesquisaPreCadastro').prop('disabled', true)
            $('#codPlanta').val(dados.codPlanta).trigger('change.select2');

            if (dados.ftOriginal) {
                exibirAvisoAlteraCancela(dados.ftOriginal, dados.tipoSolicitacao);
            }

            var p1 = carregarDesenvolvedores();
            var p2 = carregarEstilos();
            var p3 = carregarEstruturasPapelao();
            var p4 = carregarConteudosEmbalado();
            var p5 = carregarTiposProdutoPrimario();
            var p6 = carregarPaletes();
        
            $.when(p1, p2, p3, p4, p5, p6).then(function() {
                $('#idDesenvolvedor').val(dados.idDesenvolvedor).trigger('change');
                $('#codCliente').val(dados.codCliente);
                $('#nomeCliente').val(dados.nomeCliente);
                $('#codLoja').val(dados.codLoja);
                $('#tipoSolicitacao').val(dados.tipoSolicitacao).trigger('change');
                $('#referencia').val(dados.referencia);
                $('#toleranciaMaxima').val(dados.toleranciaMaxima);
                $('#toleranciaMinima').val(dados.toleranciaMinima);
                $('#estilo').val(dados.estilo).trigger('change');
                $('#observacoes').val(dados.observacoes).trigger('change');
                $('#comprimento').val(dados.comprimento || 0);
                $('#largura').val(dados.largura || 0);
                $('#altura').val(dados.altura || 0);
                $('#numeroMatriz').val(dados.numeroMatriz || 0);
                $('#inkJet').val(dados.inkJet).trigger('change');

                $('#necessitaAcompanhamento').val(dados.necessitaAcompanhamento).trigger('change');
                if (dados.necessitaAcompanhamento === 'S') {
                    $('#divNivel').show();
                    $('#nivelLotePiloto').val(dados.nivelLotePiloto).trigger('change');
                }

                $('#clienteUsaTara').val(dados.clienteUsaTara).trigger('change');
                if (dados.clienteUsaTara == '1') {
                    $('#divToleranciaPeso').show();
                    $('#toleranciaPeso').val(dados.toleranciaPeso || 0);
                }

                $('#estruturaPapelao').val(dados.estruturaPapelao).trigger('change');
                $('#gramatura').val(dados.gramatura || 0);
                $('#larguraChapa').val(dados.larguraChapa || 0);
                $('#comprimentoChapa').val(dados.comprimentoChapa || 0);
                $('#areaLiquida').val(dados.areaLiquida || 0);

                calcularAreaBlank();
                calcularPeso();

                $('#gramaturaReferencial').val(dados.gramaturaReferencial || 0);
                $('#esmagamento').val(dados.esmagamento || 0);
                $('#ectColuna').val(dados.ectColuna || 0);
                $('#mullenReferencial').val(dados.mullenReferencial || 0);
                $('#espessura').val(dados.espessura || 0);
                $('#bctCompressao').val(dados.bctCompressao || 0);
                $('#cobbInterno').val(dados.cobbInterno || 0);
                $('#cobbExterno').val(dados.cobbExterno || 0);

                $('#tipoFechamento').val(dados.tipoFechamento).trigger('change');
                $('#pecasPorUnidade').val(dados.pecasPorUnidade).trigger('change');

                $('#ehAmarrado').val(dados.ehAmarrado).trigger('change');
                if (dados.ehAmarrado === 'S') {
                    $('#divQuantidadeFitilhos').show();
                    $('#divTipoFitilho').show();
                    $('#quantidadeFitilhos').val(dados.quantidadeFitilhos).trigger('change');
                    $('#tipoFitilho').val(dados.tipoFitilho).trigger('change');
                }

                $('#ehPaletizado').val(dados.ehPaletizado).trigger('change');
                if (dados.ehPaletizado === 'S') {
                    $('#camposPaletizado').show();
                    $('#alturaMaximaPalete').val(dados.alturaMaximaPalete);
                    $('#palete').val(dados.palete).trigger('change');
                    $('#temFacesOpostas').val(dados.temFacesOpostas).trigger('change');
                    $('#temCantinho').val(dados.temCantinho).trigger('change');
                    $('#temGradeMadeira').val(dados.temGradeMadeira).trigger('change');
                    $('#temFilmeStrech').val(dados.temFilmeStrech).trigger('change');
                    $('#clienteUsaRack').val(dados.clienteUsaRack).trigger('change');
                    $('#overhang').val(dados.overhang).trigger('change');
                    $('#quantidadeEspelhos').val(dados.quantidadeEspelhos).trigger('change');
                }

                $('#conteudoEmbalado').val(dados.conteudoEmbalado).trigger('change');
                $('#consumoMensal').val(dados.consumoMensal || 0);
                $('#tipoLinhaProducaoCliente').val(dados.tipoLinhaProducaoCliente).trigger('change');
                $('#tipoProdutoPrimario').val(dados.tipoProdutoPrimario).trigger('change');
                $('#pesoPrimario').val(dados.pesoPrimario || 0);
                $('#quantidadePrimarios').val(dados.quantidadePrimarios || 0);
                $('#suportacaoPrimario').val(dados.suportacaoPrimario).trigger('change');
                $('#tempoEstocagem').val(dados.tempoEstocagem || 0);
                $('#umidadeRelativa').val(dados.umidadeRelativa || 0);
                $('#quantidadeCores').val(dados.quantidadeCores || 0);
                $('#precoKG').val(dados.precoKG || 0);
                $('#orcamento').val(dados.orcamento).trigger('change');
                $('#margemContribuicao').val(dados.margemContribuicao);
                $('#resultado').val(dados.resultado || 0);
                $('#tipoSavings').val(dados.tipoSavings).trigger('change');
                $('#valorSavingSK').val(dados.valorSavingSK || 0);
                $('#valorSavingCliente').val(dados.valorSavingCliente || 0);
                $('#valorEstimadoCliche').val(dados.valorEstimadoCliche || 0);
                $('#valorEstimadoMatriz').val(dados.valorEstimadoMatriz || 0);
                $('#informacoesAdicionais').val(dados.informacoesAdicionais);

                $('#quantidadeAmostras').val(dados.quantidadeAmostras).trigger('change');
                var qtdAmostras = parseInt(dados.quantidadeAmostras) || 0;
                if (qtdAmostras >= 1) {
                    $('#camposAmostra').show();
                    $('#amostraCep').val(dados.amostraCep).trigger('input');
                    $('#amostraResponsavel').val(dados.amostraResponsavel);
                    $('#amostraEndereco').val(dados.amostraEndereco);
                    $('#amostraBairro').val(dados.amostraBairro);
                    $('#amostraMunicipio').val(dados.amostraMunicipio);
                    $('#amostraUF').val(dados.amostraUF).trigger('change');
                    $('#amostraCNPJ').val(dados.amostraCNPJ).trigger('input');
                    $('#amostraTelefone').val(dados.amostraTelefone).trigger('input');
                    $('#amostraEmail').val(dados.amostraEmail);
                    $('#amostraObservacoes').val(dados.amostraObservacoes);
                }

                listaAcessorios = dados.acessorios || []
                listaAnexos = dados.anexos || [];
                
                renderizarListaAcessorios();
                renderizarListaAnexos();
            });

            if(modoEdicao && (statusSolicitacao === 'A' || statusSolicitacao === 'G')) {
                $('#btnSalvarEnviar').show()
            }
        }
        
        function anexarEventos() {
            $('#btnProximo').on('click', function() {
                navegarParaAba(abaAtual + 1);
            });

            $('#btnVoltar').on('click', function() { navegarParaAba(abaAtual - 1); });

            $('#btnSalvarAcessorio').on('click', function () {
                validacoes.modalAcessorio.revalidate().then(isValid => {
                    if (!isValid) return;
                    salvarAcessorio();
                });
            });

            $('#ehAmarrado').on('change', campoAmarradoAlterado);
            $('#clienteUsaTara').on('change', campoClienteUsaTaraAlterado);
            $('#codPlanta').on('change', campoPlantaAlterado);
            $('#necessitaAcompanhamento').on('change', campoNecessitaAcompanhamentoAlterado);
            $('#estruturaPapelao').on('change', campoEstruturaPapelaoAlterado);
            $('#quantidadeAmostras').on('change', campoQuantidadeAmostrasAlterado);
            $('#ehPaletizado').on('change', campoPaletizadoAlterado);
            $('#tipoSolicitacao').on('change', campoTipoSolicitacaoAlterado);

            $('#larguraChapa').on('change', calcularAreaBlank);
            $('#comprimentoChapa').on('change', calcularAreaBlank);
            $('#areaLiquida').on('change', calcularPeso);
            $('#gramatura').on('change', calcularPeso);
            
            $('#amostraCep').on('blur', buscarCep);

            let copiarFichaTimeout = null;
            $('#copiarFichaBusca').on('input', function() {
                clearTimeout(copiarFichaTimeout);
                copiarFichaTimeout = setTimeout(function() {
                    pesquisarFichaTecnica();
                }, 500);
            });

            $('#copiarFichaBusca').on('keypress', function(e) {
                if (e.which === 13) {
                    e.preventDefault();
                    clearTimeout(copiarFichaTimeout);
                    pesquisarFichaTecnica();
                }
            });
            
            $('.tab-button').on('click', function() {
                const tabNumber = parseInt($(this).data('tab'));
                if (tabNumber === abaAtual) return;
                navegarParaAba(tabNumber);
            });
        }

        async function validarFormulario(flagEnviarPCP) {
            if (modoVisualizacao) return;

            const resultados = await Promise.all([
                validacoes.aba1.revalidate(),
                validacoes.aba2.revalidate(),
                validacoes.aba3.revalidate(),
                validacoes.aba5.revalidate(),
                validacoes.aba6.revalidate(),
            ]);

            resultados.forEach((valida, index) => {
                const numeroAba = [1, 2, 3, 5, 6][index];
                if (!valida) {
                    marcarAbaComoInvalida(numeroAba);
                    exibirErrosValidacao(validacoes[`aba${numeroAba}`]);
                } else {
                    marcarAbaComoCompleta(numeroAba); 
                }
            });

            const todasValidas = resultados.every(Boolean);

            if (!todasValidas) {
                const primeiraInvalida = [1, 2, 3, 5, 6][resultados.findIndex(v => !v)];
                navegarParaAba(primeiraInvalida);
                Toast.error('Existem campos obrigatórios pendentes. Verifique as seções destacadas.');
                return;
            }

            enviarFormulario(flagEnviarPCP);
        }

        function enviarFormulario(flagEnviarPCP) {
            exibirLoading(true);

            const anexos = uploadAnexos ? uploadAnexos.getFiles() : [];

            const dadosFormulario = {
                acao: modoEdicao ? 'editar' : 'criar',
                flagEnviarPCP: flagEnviarPCP,
                codPlanta: $('#codPlanta').val(),
                idDesenvolvedor: $('#idDesenvolvedor').val(),
                codCliente: $('#codCliente').val(),
                nomeCliente: $('#nomeCliente').val(),
                codLoja: $('#codLoja').val(),
                referencia: $('#referencia').val(),
                necessitaAcompanhamento: $('#necessitaAcompanhamento').val(),
                nivelLotePiloto: $('#nivelLotePiloto').val(),
                toleranciaMaxima: $('#toleranciaMaxima').val(),
                toleranciaMinima: $('#toleranciaMinima').val(),
                estilo: $('#estilo').val(),
                observacoes: $('#observacoes').val(),
                comprimento: $('#comprimento').val(),
                largura: $('#largura').val(),
                altura: $('#altura').val(),
                numeroMatriz: $('#numeroMatriz').val(),
                clienteUsaTara: $('#clienteUsaTara').val(),
                toleranciaPeso: $('#toleranciaPeso').val(),
                estruturaPapelao: $('#estruturaPapelao').val(),
                gramatura: $('#gramatura').val(),
                larguraChapa: $('#larguraChapa').val(),
                comprimentoChapa: $('#comprimentoChapa').val(),
                areaBlank: $('#areaBlank').val(),
                areaLiquida: $('#areaLiquida').val(),
                peso: $('#peso').val(),
                inkJet: $('#inkJet').val(),
                gramaturaReferencial: $('#gramaturaReferencial').val(),
                esmagamento: $('#esmagamento').val(),
                ectColuna: $('#ectColuna').val(),
                mullenReferencial: $('#mullenReferencial').val(),
                espessura: $('#espessura').val(),
                bctCompressao: parseFloat($('#bctCompressao').val() || 0).toFixed(2),
                cobbInterno: $('#cobbInterno').val(),
                cobbExterno: $('#cobbExterno').val(),
                tipoFechamento: $('#tipoFechamento').val(),
                ehAmarrado: $('#ehAmarrado').val(),
                quantidadeFitilhos: $('#quantidadeFitilhos').val(),
                tipoFitilho: $('#tipoFitilho').val(),
                pecasPorUnidade: $('#pecasPorUnidade').val(),
                ehPaletizado: $('#ehPaletizado').val(),
                alturaMaximaPalete: $('#alturaMaximaPalete').val(),
                palete: $('#palete').val(),
                temFacesOpostas: $('#temFacesOpostas').val(),
                temCantinho: $('#temCantinho').val(),
                temGradeMadeira: $('#temGradeMadeira').val(),
                temFilmeStrech: $('#temFilmeStrech').val(),
                clienteUsaRack: $('#clienteUsaRack').val(),
                overhang: $('#overhang').val(),
                quantidadeEspelhos: $('#quantidadeEspelhos').val(),
                acessorios: JSON.stringify(listaAcessorios),
                conteudoEmbalado: $('#conteudoEmbalado').val(),
                consumoMensal: $('#consumoMensal').val(),
                tipoLinhaProducaoCliente: $('#tipoLinhaProducaoCliente').val(),
                tipoProdutoPrimario: $('#tipoProdutoPrimario').val(),
                pesoPrimario: $('#pesoPrimario').val(),
                quantidadePrimarios: $('#quantidadePrimarios').val(),
                suportacaoPrimario: $('#suportacaoPrimario').val(),
                tempoEstocagem: $('#tempoEstocagem').val(),
                umidadeRelativa: $('#umidadeRelativa').val(),
                quantidadeCores: $('#quantidadeCores').val(),
                precoKG: $('#precoKG').val(),
                orcamento: $('#orcamento').val(),
                margemContribuicao: $('#margemContribuicao').val(),
                resultado: $('#resultado').val(),
                tipoSavings: $('#tipoSavings').val(),
                valorSavingSK: $('#valorSavingSK').val(),
                valorSavingCliente: $('#valorSavingCliente').val(),
                valorEstimadoCliche: $('#valorEstimadoCliche').val(),
                valorEstimadoMatriz: $('#valorEstimadoMatriz').val(),
                informacoesAdicionais: $('#informacoesAdicionais').val(),
                quantidadeAmostras: $('#quantidadeAmostras').val(),
                amostraCep: $('#amostraCep').val()?.replace(/\D/g, ''),
                amostraResponsavel: $('#amostraResponsavel').val(),
                amostraEndereco: $('#amostraEndereco').val(),
                amostraBairro: $('#amostraBairro').val(),
                amostraMunicipio: $('#amostraMunicipio').val(),
                amostraUF: $('#amostraUF').val(),
                amostraCNPJ: $('#amostraCNPJ').val()?.replace(/\D/g, ''),
                amostraTelefone: $('#amostraTelefone').val()?.replace(/\D/g, ''),
                amostraEmail: $('#amostraEmail').val(),
                amostraObservacoes: $('#amostraObservacoes').val(),
                numeroFTCopia: $('#numeroFTCopia').val(),
                ehPreCadastro: $('#hidEhPrecadastro').val(),
                tipoSolicitacao: $('#tipoSolicitacao').val(),
                anexos: JSON.stringify(anexos),
            };
            

            if (id) {
                dadosFormulario.id = id;
            }

            $.ajax({
                url: 'SolicitacaoProjetoRepositorio.asp',
                type: 'POST',
                data: dadosFormulario,
                dataType: 'json',
                success: function(response) {
                    if (!response.sucesso) {
                        Toast.error(response.mensagem);
                        return;
                    }

                    Toast.success(response.mensagem);
                    navegarParaGrade();
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
                type: 'danger',
                confirmText: 'Confirmar',
                cancelText: 'Cancelar',
                onConfirm: () => navegarParaGrade(),
                onCancel: () => {}
            });
        }

        $(document).ready(function() {
            verificarModoEdicao();
            anexarEventos();
            navegarParaAba(1);

            if (!modoVisualizacao) {
                inicializarValidacoes();
            } else {
                aplicarModoVisualizacao();
            }

            inicializarSelect2();
            inicializarMascaras();

            renderizarListaAcessorios();
            campoQuantidadeAmostrasAlterado();

            if (modoEdicao || modoVisualizacao) {
                carregarDadosEdicao();
            }
        });
    </script>
</head>
<body>
    <div class="container my-4">
        <div class="d-flex my-4 justify-content-between align-items-center flex-wrap gap-2">
            <div class="d-flex align-items-center">
                <button type="button" onclick="navegarParaGrade()" class="btn btn-sm btn-light me-2 fs-5">
                    <i class="bi bi-arrow-left"></i>
                </button>
                <h3 class="mb-0">Solicitação de Projeto ao Desenvolvimento (SPD) #<span id="idTitulo"></span></h3>
            </div>

            <div id="divAcoes" class="page-actions flex-wrap">
                <button type="button" onclick="validarFormulario(1)" class="btn btn-sm btn-outline-primary" id="btnSalvarEnviar" style="display:none;">
                    <i class="bi bi-shield-check me-1"></i>
                    <span>Salvar e enviar para PCP</span>
                </button>

                <button type="button" onclick="validarFormulario(0)" class="btn btn-sm btn-success" id="btnSalvar">
                    <i class="bi bi-check-lg me-1"></i>
                    <span>Enviar Formulário</span>
                </button>
            </div>
        </div>

        <div id="loadingOverlay" class="position-fixed top-0 start-0 w-100 h-100 d-none justify-content-center align-items-center" style="background: rgba(255,255,255,0.8); z-index: 9999;">
            <div class="spinner-border text-primary" role="status">
                <span class="visually-hidden">Carregando...</span>
            </div>
        </div>

        <div id="avisoAlteraCancela" class="w-100 mb-4" style="display: none;"></div>

        <div id="avisoAlteracao" style="display: none; font-size: 0.85rem;" class="mb-4 alert alert-warning alert-dismissible py-2 px-3 mb-0 align-items-center" role="alert">
            <i class="bi bi-info-circle me-2"></i>
            Você possui alterações não salvas!
        </div>

        <div class="form-container">
            <div class="form-layout">
                <div class="tabs-container">
                <div class="tabs">
                    <button type="button" class="tab-button active" data-tab="1"><i class="tab-icon bi bi-box"></i><span class="tab-title">Produto</span></button>
                    <button type="button" class="tab-button" data-tab="2"><i class="tab-icon bi bi-list-ul"></i><span class="tab-title">Especificações</span></button>
                    <button type="button" class="tab-button" data-tab="3"><i class="tab-icon bi bi-brush"></i><span class="tab-title">Acabamento</span></button>
                    <button type="button" class="tab-button" data-tab="4"><i class="tab-icon bi bi-puzzle"></i><span class="tab-title">Acessórios</span></button>
                    <button type="button" class="tab-button" data-tab="5"><i class="tab-icon bi bi-archive"></i><span class="tab-title">Embalagem</span></button>
                    <button type="button" class="tab-button" data-tab="6"><i class="tab-icon bi bi-clipboard-check"></i><span class="tab-title">Amostra</span></button>
                    <button type="button" class="tab-button" data-tab="7"><i class="tab-icon bi bi-paperclip"></i><span class="tab-title">Anexos</span></button>
                </div>
                </div>

                <div class="tab-content-wrapper">
                    <form id="formularioSolicitacaoProjeto" name="formularioSolicitacaoProjeto" novalidate>
                <input type="hidden" id="hidEhPrecadastro" name="hidEhPrecadastro">
                <input type="hidden" id="hidCEP" name="hidCEP">
                <input type="hidden" id="hidEndereco" name="hidEndereco">
                <input type="hidden" id="hidBairro" name="hidBairro">
                <input type="hidden" id="hidMunicipio" name="hidMunicipio">
                <input type="hidden" id="hidUF" name="hidUF">
                <input type="hidden" id="hidTelefone" name="hidTelefone">
                <input type="hidden" id="hidCNPJ" name="hidCNPJ">
                <input type="hidden" id="hidEmail" name="hidEmail">

                <!-- Aba Produto -->
                <div class="tab-content active" data-tab-content="1">
                    <div class="section-header">
                        <div class="section-title">Produto</div>
                        <p class="section-subtitle">Dados principais da solicitação e identificação do cliente.</p>
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
                            </select>
                        </div>

                        <div class="col-md-4">
                            <label class="form-label">Buscar cliente</label>
                            <div class="input-group">
                                <span class="input-group-text py-0 px-2">
                                    <div class="form-check form-switch mb-0 d-flex align-items-center gap-2">
                                        <input class="form-check-input" type="checkbox" role="switch" id="flagPesquisaPreCadastro" style="cursor:pointer">
                                        <label class="form-check-label small text-muted" for="flagPesquisaPreCadastro">
                                            PRÉ-CADASTRO
                                        </label>
                                    </div>
                                </span>
                                <button id="btnBuscarCliente" type="button" class="flex-grow-1 btn btn-sm btn-secondary" onclick="pesquisarCliente()">
                                    Buscar <i class="bi bi-box-arrow-up-right"></i>
                                </button>
                            </div>
                        </div>

                        <div class="col-md-4">
                            <label for="nomeCliente" class="form-label">Nome do Cliente <span class="required-field">*</span></label>
                            <input type="text" class="form-control form-control-sm" id="nomeCliente" name="nomeCliente" disabled>
                        </div>

                        <div class="col-md-2">
                            <label for="codCliente" class="form-label">Cod. Cliente <span class="required-field">*</span></label>
                            <input type="text" class="form-control form-control-sm" id="codCliente" name="codCliente" disabled>
                        </div>

                        <div class="col-md-2">
                            <label for="codLoja" class="form-label">Loja <span class="required-field">*</span></label>
                            <input type="text" class="form-control form-control-sm" id="codLoja" name="codLoja" disabled>
                        </div>

                        <div class="col-lg-4 col-md-6">
                            <label for="tipoSolicitacao" class="form-label">Tipo de Solicitação <span class="required-field">*</span></label>
                            <select class="form-select form-select-sm select2" name="tipoSolicitacao" id="tipoSolicitacao">
                                <option value="">Selecione o tipo de solicitação</option>
                                <option value="1">Projetar/alterar exatamente conforme desenho ou amostra enviado ao Desenvolvimento</option>
                                <option value="2">Projetar/alterar exatamente conforme especificações identificadas nesta Solicitação</option>
                                <option value="3">Alterar ou Cancelar uma Ficha Técnica</option>
                                <option value="6">Novo Projeto / Primário</option>
                            </select>
                        </div>

                        <div class="col-lg-4 col-md-6">
                            <label for="numeroFTCopia" class="form-label" id="labelNumeroFTCopia"><span id="textoLabelFTCopia">Copiar de FT</span><span id="asteriscoFTCopia" class="required-field" style="display:none;"> *</span></label>
                            <div class="input-group">
                                <button id="btnCopiarFicha" type="button" class="btn btn-sm btn-secondary" onclick="abrirModalCopiarFicha()">
                                    Buscar <i class="bi bi-box-arrow-up-right"></i>
                                </button>
                                <input type="text" class="form-control form-control-sm" id="numeroFTCopia" name="numeroFTCopia">
                            </div>
                        </div>

                        <div class="col-lg-4 col-md-6">
                            <label for="idDesenvolvedor" class="form-label">Técnico/Desenvolvedor <span class="required-field">*</span></label>
                            <select class="form-select form-select-sm select2" id="idDesenvolvedor" name="idDesenvolvedor"><option value="">Selecione o técnico...</option></select>
                        </div>

                        <div class="col-lg-4 col-md-6">
                            <label for="referencia" class="form-label">Referência <span class="required-field">*</span></label>
                            <input type="text" class="form-control form-control-sm" id="referencia" name="referencia">
                        </div>

                        <div class="col-lg-4 col-md-6">
                            <label for="necessitaAcompanhamento" class="form-label">Necessita Acompanhamento <span class="required-field">*</span></label>
                            <select class="form-select form-select-sm select2" id="necessitaAcompanhamento" name="necessitaAcompanhamento">
                                <option value="">Selecione...</option>
                                <option value="S">Sim</option>
                                <option value="N">Não</option>
                            </select>
                        </div>

                        <div class="col-lg-4 col-md-6" id="divNivel" style="display: none;">
                            <label for="nivelLotePiloto" class="form-label">Nível <span class="required-field">*</span></label>
                            <select class="form-select form-select-sm select2" id="nivelLotePiloto" name="nivelLotePiloto">
                                <option value="">Selecione...</option>
                                <option value="1">Nível 1</option>
                                <option value="2">Nível 2</option>
                            </select>
                        </div>

                        <div class="col-lg-4 col-md-6">
                            <label for="toleranciaMaxima" class="form-label">Tolerância Máxima (%) <span class="required-field">*</span></label>
                            <input type="number" class="form-control form-control-sm" id="toleranciaMaxima" name="toleranciaMaxima" min="0" max="100">
                        </div>

                        <div class="col-lg-4 col-md-6">
                            <label for="toleranciaMinima" class="form-label">Tolerância Mínima (%) <span class="required-field">*</span></label>
                            <input type="number" class="form-control form-control-sm" id="toleranciaMinima" name="toleranciaMinima" min="0" max="100">
                        </div>

                        <div class="col-lg-4 col-md-6">
                            <div class="d-flex align-items-center gap-2">
                                <label for="estilo" class="form-label mb-0">Estilo <span class="required-field">*</span></label>
                                <a href="javascript:void(0)" onclick="abrirPdfFefco()" class="btn btn-outline-primary btn-help" title="Ver catálogo de estilos FEFCO">
                                    <i class="bi bi-file-earmark-pdf"></i> Catálogo FEFCO
                                </a>
                            </div>
                            <select class="form-select form-select-sm select2" id="estilo" name="estilo">
                                <option value="">Selecione o estilo...</option>
                            </select>
                        </div>

                        <div class="col-lg-4 col-md-6">
                            <label for="observacoes" class="form-label">Observações</label>
                            <select class="form-select form-select-sm select2" id="observacoes" name="observacoes">
                                <option value="">Selecione...</option>
                                <option value="F">Com Furos</option>
                                <option value="C">Com Chanfros</option>
                                <option value="E">Com Furos/Chanfros</option>
                                <option value="D">Com Furo Alça</option>
                                <option value="A">Caixa com Display</option>
                            </select>
                        </div>

                        <div class="col-lg-4 col-md-6">
                            <label for="comprimento" class="form-label">Comprimento (mm) <span class="required-field">*</span></label>
                            <input type="number" class="form-control form-control-sm" id="comprimento" name="comprimento" min="0">
                        </div>

                        <div class="col-lg-4 col-md-6">
                            <label for="largura" class="form-label">Largura (mm) <span class="required-field">*</span></label>
                            <input type="number" class="form-control form-control-sm" id="largura" name="largura" min="0">
                        </div>

                        <div class="col-lg-4 col-md-6">
                            <label for="altura" class="form-label">Altura (mm) <span class="required-field">*</span></label>
                            <input type="number" class="form-control form-control-sm" id="altura" name="altura" min="0">
                        </div>

                        <div class="col-lg-4 col-md-6">
                            <label for="numeroMatriz" class="form-label">Número da Matriz (Faca Rotativa)<span class="required-field">*</span></label>
                            <input type="number" class="form-control form-control-sm" id="numeroMatriz" name="numeroMatriz" min="0">
                        </div>

                        <div class="col-lg-4 col-md-6">
                            <label for="clienteUsaTara" class="form-label">Cliente usa Tara <span class="required-field">*</span></label>
                            <select class="form-select form-select-sm select2" id="clienteUsaTara" name="clienteUsaTara">
                                <option value="">Selecione...</option>
                                <option value="1">Sim</option>
                                <option value="2">Não</option>
                            </select>
                        </div>

                        <div class="col-lg-4 col-md-6" id="divToleranciaPeso" style="display: none;">
                            <label for="toleranciaPeso" class="form-label">Tolerância de Peso ( ± %) <span class="required-field">*</span></label>
                            <input type="number" class="form-control form-control-sm" id="toleranciaPeso" name="toleranciaPeso" min="0" max="100">
                        </div>

                        <div class="col-lg-4 col-md-6">
                            <label for="estruturaPapelao" class="form-label">Estrutura de Papelão Ondulado <span class="required-field">*</span></label>
                            <select class="form-select form-select-sm select2" id="estruturaPapelao" name="estruturaPapelao">
                                <option value="">Selecione a estrutura...</option>
                            </select>
                        </div>

                        <div class="col-lg-4 col-md-6">
                            <label for="gramatura" class="form-label">Gramatura <span class="required-field">*</span></label>
                            <input disabled type="number" class="form-control form-control-sm" id="gramatura" name="gramatura" min="0">
                        </div>

                        <div class="col-lg-4 col-md-6">
                            <label for="larguraChapa" class="form-label">Largura da Chapa (mm) <span class="required-field">*</span></label>
                            <input type="number" class="form-control form-control-sm" id="larguraChapa" name="larguraChapa" min="0">
                        </div>

                        <div class="col-lg-4 col-md-6">
                            <label for="comprimentoChapa" class="form-label">Comprimento da Chapa (mm) <span class="required-field">*</span></label>
                            <input type="number" class="form-control form-control-sm" id="comprimentoChapa" name="comprimentoChapa" min="0">
                        </div>

                        <div class="col-lg-4 col-md-6">
                            <label for="areaBlank" class="form-label">Área Blank (m²) <span class="required-field">*</span></label>
                            <input disabled type="number" class="form-control form-control-sm" id="areaBlank" name="areaBlank" min="0">
                        </div>

                        <div class="col-lg-4 col-md-6">
                            <label for="areaLiquida" class="form-label">Área Líquida (m²) <span class="required-field">*</span></label>
                            <input type="number" class="form-control form-control-sm" id="areaLiquida" name="areaLiquida" min="0">
                        </div>

                        <div class="col-lg-4 col-md-6">
                            <label for="peso" class="form-label">Peso (g) <span class="required-field">*</span></label>
                            <input disabled type="number" class="form-control form-control-sm" id="peso" name="peso" min="0">
                        </div>

                        <div class="col-lg-4 col-md-6">
                            <label for="inkJet" class="form-label">Ink Jet <span class="required-field">*</span></label>
                            <select class="form-select form-select-sm select2" id="inkJet" name="inkJet">
                                <option value="">Selecione...</option>
                                <option value="S">Sim</option>
                                <option value="N">Não</option>
                            </select>
                        </div>
                    </div>
                </div>

                <!-- Aba Especificações -->
                <div class="tab-content" data-tab-content="2">
                    <div class="section-header">
                        <div class="section-title">Especificações</div>
                        <p class="section-subtitle">Informações técnicas e parâmetros estruturais do produto.</p>
                    </div>
                    <div class="row g-3">
                        <div class="col-lg-4 col-md-6">
                            <label for="gramaturaReferencial" class="form-label">Gramatura Referencial (g/m²) <span class="required-field">*</span></label>
                            <input type="number" class="form-control form-control-sm" id="gramaturaReferencial" name="gramaturaReferencial" min="0" step="1">
                        </div>

                        <div class="col-lg-4 col-md-6">
                            <label for="esmagamento" class="form-label">Esmagamento (kgf/cm²) <span class="required-field">*</span></label>
                            <input type="number" class="form-control form-control-sm" id="esmagamento" name="esmagamento" min="0" step="1">
                        </div>

                        <div class="col-lg-4 col-md-6">
                            <label for="ectColuna" class="form-label">ECT/Coluna (kgf/cm) <span class="required-field">*</span></label>
                            <input type="number" class="form-control form-control-sm" id="ectColuna" name="ectColuna" min="0" step="1">
                        </div>

                        <div class="col-lg-4 col-md-6">
                            <label for="mullenReferencial" class="form-label">Mullen Referencial (kgf/cm²) <span class="required-field">*</span></label>
                            <input type="number" class="form-control form-control-sm" id="mullenReferencial" name="mullenReferencial" min="0" step="1">
                        </div>

                        <div class="col-lg-4 col-md-6">
                            <label for="espessura" class="form-label">Espessura (mm) <span class="required-field">*</span></label>
                            <input type="number" class="form-control form-control-sm" id="espessura" name="espessura" min="0" step="1">
                        </div>

                        <div class="col-lg-4 col-md-6">
                            <label for="bctCompressao" class="form-label">BCT/Compressão (kgf) <span class="required-field">*</span></label>
                            <input type="number" class="form-control form-control-sm" id="bctCompressao" name="bctCompressao" min="0" step="1">
                        </div>

                        <div class="col-lg-4 col-md-6">
                            <label for="cobbInterno" class="form-label">COBB Interno (g/m²H₂O) <span class="required-field">*</span></label>
                            <input type="number" class="form-control form-control-sm" id="cobbInterno" name="cobbInterno" min="0" step="1">
                        </div>

                        <div class="col-lg-4 col-md-6">
                            <label for="cobbExterno" class="form-label">COBB Externo (g/m²H₂O) <span class="required-field">*</span></label>
                            <input type="number" class="form-control form-control-sm" id="cobbExterno" name="cobbExterno" min="0" step="1">
                        </div>
                    </div>

                    <div class="card bg-light p-2 mt-4">
                        <div class="d-flex gap-2 align-items-center">
                            <code>1 kgf = 9,80665 N</code>
                            <span class="text-muted">Para converter de kgf para Newton, multiplique por 9,80665</span>
                        </div>
                        <div class="d-flex gap-2 align-items-center">
                            <code>1 lbf/in² = 0,07031 kgf/cm²</code>
                            <span class="text-muted">Para converter de lbf/in² para kgf/cm², multiplique por 0,07031</span>
                        </div>
                        <div class="d-flex gap-2 align-items-center">
                            <code>1 in = 25,4 mm</code>
                            <span class="text-muted">Para converter de polegadas para mm, multiplique por 25,4</span>
                        </div>
                    </div>
                </div>

                <!-- Aba Acabamento -->
                <div class="tab-content" data-tab-content="3">
                    <div class="section-header">
                        <div class="section-title">Acabamento</div>
                        <p class="section-subtitle">Dados de fechamento, amarração e paletização.</p>
                    </div>
                    <div class="row g-3">
                        <div class="col-lg-4 col-md-6">
                            <label for="tipoFechamento" class="form-label">Fechamento <span class="required-field">*</span></label>
                            <select class="form-select form-select-sm select2" id="tipoFechamento" name="tipoFechamento">
                                <option value="">Selecione...</option>
                                <option value="C">C - Colado</option>
                                <option value="G">G - Grampeado</option>
                                <option value="S">S - Sem Fechamento</option>
                            </select>
                        </div>

                        <div class="col-lg-4 col-md-6">
                            <label for="ehAmarrado" class="form-label">Amarrado <span class="required-field">*</span></label>
                            <select class="form-select form-select-sm select2" id="ehAmarrado" name="ehAmarrado">
                                <option value="">Selecione...</option>
                                <option value="S">F - Fitilhado</option>
                                <option value="N">S - Sem Amarrar</option>
                                <option value="C">C - Caixa Master</option>
                            </select>
                        </div>

                        <div class="col-lg-4 col-md-6" id="divQuantidadeFitilhos" style="display: none;">
                            <label for="quantidadeFitilhos" class="form-label">Quantidade de Fitilhos <span class="required-field">*</span></label>
                            <select class="form-select form-select-sm select2" id="quantidadeFitilhos" name="quantidadeFitilhos">
                                <option value="">Selecione...</option>
                                <option value="1">1</option>
                                <option value="2">2</option>
                            </select>
                        </div>

                        <div class="col-lg-4 col-md-6" id="divTipoFitilho" style="display: none;">
                            <label for="tipoFitilho" class="form-label">Fitilhos na Largura ou Comprimento <span class="required-field">*</span></label>
                            <select class="form-select form-select-sm select2" id="tipoFitilho" name="tipoFitilho">
                                <option value="">Selecione...</option>
                                <option value="L">L - Na largura</option>
                                <option value="C">C - No comprimento</option>
                            </select>
                        </div>

                        <div class="col-lg-4 col-md-6">
                            <label for="pecasPorUnidade" class="form-label">Pacotes de <span class="required-field">*</span></label>
                            <select class="form-select form-select-sm select2" id="pecasPorUnidade" name="pecasPorUnidade">
                                <option value="">Selecione...</option>
                                <option value="5">5</option>
                                <option value="10">10</option>
                                <option value="15">15</option>
                                <option value="20">20</option>
                                <option value="25">25</option>
                                <option value="30">30</option>
                                <option value="35">35</option>
                                <option value="40">40</option>
                                <option value="45">45</option>
                                <option value="50">50</option>
                            </select>
                        </div>

                        <div class="col-lg-4 col-md-6">
                            <label for="ehPaletizado" class="form-label">Paletizado <span class="required-field">*</span></label>
                            <select class="form-select form-select-sm select2" id="ehPaletizado" name="ehPaletizado">
                                <option value="">Selecione...</option>
                                <option value="S">Sim</option>
                                <option value="N">Não</option>
                            </select>
                        </div>

                        <div id="camposPaletizado" style="display: none;" class="mt-3">
                            <div class="row g-3">
                                <div class="col-lg-4 col-md-6">
                                    <label for="palete" class="form-label">
                                        Palete <span class="required-field">*</span>
                                    </label>
                                    <select class="form-select form-select-sm select2" id="palete">
                                        <option value="">Selecione o palete...</option>
                                    </select>
                                </div>
                                <div class="col-lg-4 col-md-6">
                                    <label for="alturaMaximaPalete" class="form-label">Altura Máxima do Palete (mm) <span class="required-field">*</span></label>
                                    <input type="number" class="form-control form-control-sm" id="alturaMaximaPalete" name="alturaMaximaPalete" min="0">
                                </div>
                                <div class="col-lg-4 col-md-6">
                                    <label for="temFacesOpostas" class="form-label">Faces Opostas <span class="required-field">*</span></label>
                                    <select class="form-select form-select-sm select2" id="temFacesOpostas" name="temFacesOpostas">
                                        <option value="">Selecione...</option>
                                        <option value="S">Sim</option>
                                        <option value="N">Não</option>
                                    </select>
                                </div>
                                <div class="col-lg-4 col-md-6">
                                    <label for="temCantinho" class="form-label">Cantinho <span class="required-field">*</span></label>
                                    <select class="form-select form-select-sm select2" id="temCantinho" name="temCantinho">
                                        <option value="">Selecione...</option>
                                        <option value="S">Sim</option>
                                        <option value="N">Não</option>
                                    </select>
                                </div>
                                <div class="col-lg-4 col-md-6">
                                    <label for="temGradeMadeira" class="form-label">Grade de Madeira <span class="required-field">*</span></label>
                                    <select class="form-select form-select-sm select2" id="temGradeMadeira" name="temGradeMadeira">
                                        <option value="">Selecione...</option>
                                        <option value="S">Sim</option>
                                        <option value="N">Não</option>
                                    </select>
                                </div>
                                <div class="col-lg-4 col-md-6">
                                    <label for="temFilmeStrech" class="form-label">Filme Strech <span class="required-field">*</span></label>
                                    <select class="form-select form-select-sm select2" id="temFilmeStrech" name="temFilmeStrech">
                                        <option value="">Selecione...</option>
                                        <option value="S">Sim</option>
                                        <option value="N">Não</option>
                                    </select>
                                </div>
                                <div class="col-lg-4 col-md-6">
                                    <label for="clienteUsaRack" class="form-label">Cliente usa Rack <span class="required-field">*</span></label>
                                    <select class="form-select form-select-sm select2" id="clienteUsaRack" name="clienteUsaRack">
                                        <option value="">Selecione...</option>
                                        <option value="S">Sim</option>
                                        <option value="N">Não</option>
                                    </select>
                                </div>
                                <div class="col-lg-4 col-md-6">
                                    <label for="overhang" class="form-label">Cliente aceita caixa para fora do palete (overhang) <span class="required-field">*</span></label>
                                    <select class="form-select form-select-sm select2" id="overhang" name="overhang">
                                        <option value="">Selecione...</option>
                                        <option value="S">Sim</option>
                                        <option value="N">Não</option>
                                    </select>
                                </div>
                                <div class="col-lg-4 col-md-6">
                                    <label for="quantidadeEspelhos" class="form-label">Espelho <span class="required-field">*</span></label>
                                    <select class="form-select form-select-sm select2" id="quantidadeEspelhos" name="quantidadeEspelhos">
                                        <option value="">Selecione...</option>
                                        <option value="0">0</option>
                                        <option value="1">1</option>
                                        <option value="2">2</option>
                                        <option value="3">3</option>
                                        <option value="4">4</option>
                                    </select>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>

                <!-- Aba Acessórios -->
                <div class="tab-content" data-tab-content="4">
                    <div class="section-header d-flex justify-content-between align-items-center">
                        <div>
                            <div class="section-title">
                                Acessórios <span class="badge bg-secondary ms-1" id="contadorAcessorios">0</span>
                            </div>
                            <p class="section-subtitle">Lista de Acessórios.</p>
                        </div>

                        <button type="button" class="btn btn-sm btn-primary" onclick="abrirModalAcessorio()">
                            <i class="bi bi-plus-lg me-1"></i>Adicionar
                        </button>
                    </div>

                    <div id="acessoriosListaContainer">
                    </div>
                </div>

                <!-- Aba Embalagem -->
                <div class="tab-content" data-tab-content="5">
                    <div class="section-header">
                        <div class="section-title">Embalagem</div>
                        <p class="section-subtitle">Dados de conteúdo, consumo, produção, orçamento e savings.</p>
                    </div>
                    <div class="row g-3">
                        <div class="col-lg-4 col-md-6">
                            <label for="conteudoEmbalado" class="form-label">Conteúdo a ser Embalado <span class="required-field">*</span></label>
                            <select class="form-select form-select-sm select2" id="conteudoEmbalado" name="conteudoEmbalado">
                                <option value="">Selecione o conteúdo...</option>
                            </select>
                        </div>

                        <div class="col-lg-4 col-md-6">
                            <label for="consumoMensal" class="form-label">Consumo Mensal (unidades) <span class="required-field">*</span></label>
                            <input type="number" class="form-control form-control-sm" id="consumoMensal" name="consumoMensal" min="0" step="1">
                        </div>

                        <div class="col-lg-4 col-md-6">
                            <label for="tipoLinhaProducaoCliente" class="form-label">Linha de Produção do Cliente <span class="required-field">*</span></label>
                            <select class="form-select form-select-sm select2" id="tipoLinhaProducaoCliente" name="tipoLinhaProducaoCliente">
                                <option value="">Selecione...</option>
                                <option value="M">Manual</option>
                                <option value="A">Automática</option>
                            </select>
                        </div>

                        <div class="col-lg-4 col-md-6">
                            <div class="d-flex align-items-center gap-2">
                                <label for="tipoProdutoPrimario" class="form-label mb-0">Tipo de Produto Primário <span class="required-field">*</span></label>
                                <a href="javascript:void(0)" onclick="abrirModalTiposProduto()" class="btn btn-outline-primary btn-help" title="Ver tipos de produto">
                                    <i class="bi bi-image"></i> Ver Tipos
                                </a>
                            </div>
                            <select class="form-select form-select-sm select2" id="tipoProdutoPrimario" name="tipoProdutoPrimario">
                                <option value="">Selecione o tipo...</option>
                            </select>
                        </div>

                        <div class="col-lg-4 col-md-6">
                            <label for="pesoPrimario" class="form-label">Peso do Primário (kg) <span class="required-field">*</span></label>
                            <input type="number" class="form-control form-control-sm" id="pesoPrimario" name="pesoPrimario" min="0" step="1">
                        </div>

                        <div class="col-lg-4 col-md-6">
                            <label for="quantidadePrimarios" class="form-label">Quantidade de Primários <span class="required-field">*</span></label>
                            <input type="number" class="form-control form-control-sm" id="quantidadePrimarios" name="quantidadePrimarios" min="0" step="1">
                        </div>

                        <div class="col-lg-4 col-md-6">
                            <label for="suportacaoPrimario" class="form-label">Suportação do Primário <span class="required-field">*</span></label>
                            <select class="form-select form-select-sm select2" id="suportacaoPrimario" name="suportacaoPrimario">
                                <option value="">Selecione...</option>
                                <option value="2">Suportante</option>
                                <option value="1">Semi Suportante</option>
                                <option value="0">Não Suportante</option>
                            </select>
                        </div>

                        <div class="col-lg-4 col-md-6">
                            <label for="tempoEstocagem" class="form-label">Tempo de Estocagem (meses) <span class="required-field">*</span></label>
                            <input type="number" class="form-control form-control-sm" id="tempoEstocagem" name="tempoEstocagem" min="0" step="1">
                        </div>

                        <div class="col-lg-4 col-md-6">
                            <label for="umidadeRelativa" class="form-label">Umidade Relativa (%) <span class="required-field">*</span></label>
                            <input type="number" class="form-control form-control-sm" id="umidadeRelativa" name="umidadeRelativa" min="0" max="100" step="1">
                        </div>

                        <div class="col-lg-4 col-md-6">
                            <label for="quantidadeCores" class="form-label">Quantidade de Cores <span class="required-field">*</span></label>
                            <input type="number" class="form-control form-control-sm" id="quantidadeCores" name="quantidadeCores" min="0" step="1">
                        </div>

                        <div class="col-lg-4 col-md-6">
                            <label for="precoKG" class="form-label">Preço KG (R$) <span class="required-field">*</span></label>
                            <input type="number" class="form-control form-control-sm" id="precoKG" name="precoKG" min="0" step="1">
                        </div>

                        <div class="col-lg-4 col-md-6">
                            <label for="orcamento" class="form-label">Orçamento</label>
                            <select class="form-select form-select-sm select2" id="orcamento" name="orcamento">
                                <option value="">Selecione...</option>
                                <option value="1">Sim</option>
                                <option value="0">Não</option>
                            </select>
                        </div>

                        <div class="col-lg-4 col-md-6">
                            <label for="margemContribuicao" class="form-label">Margem de Contribuição (%) <span class="required-field">*</span></label>
                            <input type="number" class="form-control form-control-sm" id="margemContribuicao" name="margemContribuicao" min="0" max="100" step="1">
                        </div>

                        <div class="col-lg-4 col-md-6">
                            <label for="resultado" class="form-label">Resultado (%) <span class="required-field">*</span></label>
                            <input type="number" class="form-control form-control-sm" id="resultado" name="resultado" min="0" max="100" step="1">
                        </div>

                        <div class="col-lg-4 col-md-6">
                            <label for="tipoSavings" class="form-label">Savings <span class="required-field">*</span></label>
                            <select class="form-select form-select-sm select2" id="tipoSavings" name="tipoSavings">
                                <option value="">Selecione...</option>
                                <option value="K">Saving SK</option>
                                <option value="C">Saving Cliente</option>
                                <option value="A">Ambos</option>
                            </select>
                        </div>

                        <div class="col-lg-4 col-md-6">
                            <label for="valorSavingSK" class="form-label">Valor Saving SK (R$) <span class="required-field">*</span></label>
                            <input type="number" class="form-control form-control-sm" id="valorSavingSK" name="valorSavingSK" min="0" step="1">
                        </div>

                        <div class="col-lg-4 col-md-6">
                            <label for="valorSavingCliente" class="form-label">Valor Saving Cliente (R$) <span class="required-field">*</span></label>
                            <input type="number" class="form-control form-control-sm" id="valorSavingCliente" name="valorSavingCliente" min="0" step="1">
                        </div>

                        <div class="col-lg-4 col-md-6">
                            <label for="valorEstimadoCliche" class="form-label">Valor Estimado Clichê (R$) <span class="required-field">*</span></label>
                            <input type="number" class="form-control form-control-sm" id="valorEstimadoCliche" name="valorEstimadoCliche" min="0" step="1">
                        </div>

                        <div class="col-lg-4 col-md-6">
                            <label for="valorEstimadoMatriz" class="form-label">Valor Estimado Matriz (R$) <span class="required-field">*</span></label>
                            <input type="number" class="form-control form-control-sm" id="valorEstimadoMatriz" name="valorEstimadoMatriz" min="0" step="1">
                        </div>

                        <div class="col-12">
                            <label for="informacoesAdicionais" class="form-label">Informações Adicionais</label>
                            <textarea class="form-control form-control-sm" id="informacoesAdicionais" name="informacoesAdicionais" rows="4"></textarea>
                        </div>
                    </div>
                </div>

                <!-- Aba Amostra -->
                <div class="tab-content" data-tab-content="6">
                    <div class="section-header">
                        <div class="section-title">Amostra</div>
                        <p class="section-subtitle">Informações de quantidade e entrega das amostras.</p>
                    </div>
                    <div class="row g-3">
                        <div class="col-lg-4 col-md-6">
                            <label for="quantidadeAmostras" class="form-label">Quantidade de Amostras <span class="required-field">*</span></label>
                            <select class="form-select form-select-sm select2" id="quantidadeAmostras" name="quantidadeAmostras">
                                <option value="0">0 - Não necessita amostra</option>
                                <option value="1">1</option>
                                <option value="2">2</option>
                            </select>
                        </div>
                    </div>

                    <div id="camposAmostra" style="display: none;" class="mt-3">
                        <div class="row g-3">
                            <div class="col-lg-3 col-md-4">
                                <label for="amostraCep" class="form-label">CEP <span class="required-field">*</span></label>
                                <input type="text" 
                                    class="form-control form-control-sm" 
                                    id="amostraCep" 
                                    name="amostraCep" 
                                    placeholder="00000-000"
                                    data-mask="00000-000"
                                    maxlength="9">
                                <small class="text-muted">Digite o CEP para buscar o endereço</small>
                            </div>

                            <div class="col-lg-5 col-md-8">
                                <label for="amostraResponsavel" class="form-label">Aos Cuidados <span class="required-field">*</span></label>
                                <input type="text" 
                                    class="form-control form-control-sm" 
                                    id="amostraResponsavel" 
                                    name="amostraResponsavel"
                                    placeholder="Nome do responsável pelo recebimento">
                            </div>

                            <div class="col-lg-4 col-md-12">
                                <label for="amostraTelefone" class="form-label">Telefone <span class="required-field">*</span></label>
                                <input type="text" 
                                    class="form-control form-control-sm" 
                                    id="amostraTelefone" 
                                    name="amostraTelefone"
                                    placeholder="(00) 00000-0000"
                                    maxlength="15">
                            </div>

                            <div class="col-12">
                                <label for="amostraEndereco" class="form-label">Endereço de Entrega <span class="required-field">*</span></label>
                                <input type="text" 
                                    class="form-control form-control-sm" 
                                    id="amostraEndereco" 
                                    name="amostraEndereco"
                                    placeholder="Rua, Avenida, número, complemento...">
                            </div>

                            <div class="col-lg-4 col-md-6">
                                <label for="amostraBairro" class="form-label">Bairro <span class="required-field">*</span></label>
                                <input type="text" 
                                    class="form-control form-control-sm" 
                                    id="amostraBairro" 
                                    name="amostraBairro"
                                    placeholder="Bairro">
                            </div>

                            <div class="col-lg-4 col-md-6">
                                <label for="amostraMunicipio" class="form-label">Município <span class="required-field">*</span></label>
                                <input type="text" 
                                    class="form-control form-control-sm" 
                                    id="amostraMunicipio" 
                                    name="amostraMunicipio"
                                    placeholder="Município">
                            </div>

                            <div class="col-lg-4 col-md-4">
                                <label for="amostraUF" class="form-label">UF <span class="required-field">*</span></label>
                                <select class="form-select form-select-sm select2" id="amostraUF" name="amostraUF">
                                    <option value="">UF</option>
                                    <option value="AC">AC</option>
                                    <option value="AL">AL</option>
                                    <option value="AP">AP</option>
                                    <option value="AM">AM</option>
                                    <option value="BA">BA</option>
                                    <option value="CE">CE</option>
                                    <option value="DF">DF</option>
                                    <option value="ES">ES</option>
                                    <option value="GO">GO</option>
                                    <option value="MA">MA</option>
                                    <option value="MT">MT</option>
                                    <option value="MS">MS</option>
                                    <option value="MG">MG</option>
                                    <option value="PA">PA</option>
                                    <option value="PB">PB</option>
                                    <option value="PR">PR</option>
                                    <option value="PE">PE</option>
                                    <option value="PI">PI</option>
                                    <option value="RJ">RJ</option>
                                    <option value="RN">RN</option>
                                    <option value="RS">RS</option>
                                    <option value="RO">RO</option>
                                    <option value="RR">RR</option>
                                    <option value="SC">SC</option>
                                    <option value="SP">SP</option>
                                    <option value="SE">SE</option>
                                    <option value="TO">TO</option>
                                </select>
                            </div>

                            <div class="col-lg-6 col-md-8">
                                <label for="amostraCNPJ" class="form-label">CNPJ <span class="required-field">*</span></label>
                                <input type="text" 
                                    class="form-control form-control-sm" 
                                    id="amostraCNPJ" 
                                    name="amostraCNPJ"
                                    placeholder="00.000.000/0000-00"
                                    maxlength="15">
                            </div>

                            <div class="col-lg-6 col-md-12">
                                <label for="amostraEmail" class="form-label">E-mail <span class="required-field">*</span></label>
                                <input type="email" 
                                    class="form-control form-control-sm" 
                                    id="amostraEmail" 
                                    name="amostraEmail"
                                    placeholder="email@exemplo.com.br">
                            </div>

                            <div class="col-12">
                                <label for="amostraObservacoes" class="form-label">Observações</label>
                                <textarea class="form-control form-control-sm" 
                                            id="amostraObservacoes" 
                                            name="amostraObservacoes" 
                                            rows="3"
                                            placeholder="Informações adicionais sobre a entrega da amostra..."></textarea>
                            </div>
                        </div>
                    </div>
                </div>

                <!-- Aba Dados de Anexos -->
                <div class="tab-content" data-tab-content="7">
                    <div class="section-header">
                        <div class="section-title">Anexos</div>
                        <p class="section-subtitle">Evidências, documentos e arquivos relacionados à solicitação.</p>
                    </div>

                    <div class="row g-3">
                        <file-upload id="uploadAnexos" file-code="SPD" max-files="10" multiple></file-upload>
                        <div id="anexosListaContainer"></div>
                    </div>
                </div>

                <div class="form-navigation">
                    <div class="d-flex justify-content-end align-items-center gap-2">
                        <button type="button" class="btn btn-sm btn-outline-secondary" id="btnVoltar">
                            <i class="bi bi-arrow-left me-1"></i> 
                            <span>Voltar</span>
                        </button>
                        <button type="button" class="btn btn-sm btn-primary" id="btnProximo">
                            <span>Próximo</span> 
                            <i class="bi bi-arrow-right ms-1"></i>
                        </button>
                    </div>
                </div>
                    </form>
                </div>
            </div>
        </div>
    </div>

    <!-- Modal Copiar Ficha Técnica -->
    <div class="modal fade" id="modalCopiarFicha" tabindex="-1" aria-labelledby="modalCopiarFichaLabel" aria-hidden="true">
        <div class="modal-dialog modal-xl modal-dialog-centered modal-dialog-scrollable">
            <div class="modal-content">
                <div class="modal-header">
                    <h5 class="modal-title" id="modalCopiarFichaLabel">
                        <i class="bi bi-copy me-2"></i>Copiar de uma Ficha Técnica
                    </h5>
                    <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Fechar"></button>
                </div>
                <div class="modal-body">
                    <div class="mb-3">
                        <label for="copiarFichaBusca" class="form-label">Pesquisar Ficha Técnica</label>
                        <div class="input-group">
                            <span class="input-group-text"><i class="bi bi-search"></i></span>
                            <input type="text" 
                                   class="form-control form-control-sm" 
                                   id="copiarFichaBusca" 
                                   placeholder="Digite o número da ficha ou referência para buscar..."
                                   autocomplete="off">
                        </div>
                    </div>

                    <div id="copiarFichaResultados" style="max-height: 400px; overflow-y: auto;">
                    </div>
                </div>
                <div class="modal-footer">
                    <button type="button" class="btn btn-sm btn-outline-secondary" data-bs-dismiss="modal">
                        Cancelar
                    </button>
                    <button type="button" class="btn btn-sm btn-primary" id="btnConfirmarCopiarFicha" onclick="confirmarCopiarFicha()" disabled>
                        <i class="bi bi-check-lg me-1"></i>Copiar Dados
                    </button>
                </div>
            </div>
        </div>
    </div>

    <!-- Modal de Acessório -->
    <div class="modal fade" id="modalAcessorio" tabindex="-1" aria-labelledby="modalAcessorioLabel" aria-hidden="true">
        <div class="modal-dialog modal-lg">
            <div class="modal-content">
                <div class="modal-header">
                    <h5 class="modal-title" id="modalAcessorioTitulo">
                        <i class="bi bi-puzzle me-2"></i>Adicionar Acessório
                    </h5>
                    <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Fechar"></button>
                </div>
                <div class="modal-body">
                    <form id="formularioAcessorio" name="formularioAcessorio" novalidate>
                        <div class="row g-3">
                            <div class="col-lg-4 col-md-6">
                                <label for="acessorioQuantidade" class="form-label">
                                    Quantidade <span class="required-field">*</span>
                                </label>
                                <input type="number" class="form-control form-control-sm" id="acessorioQuantidade" min="1">
                            </div>

                            <div class="col-lg-4 col-md-6">
                                <label for="acessorioEstilo" class="form-label">
                                    Estilo <span class="required-field">*</span>
                                </label>
                                <select class="form-select form-select-sm select2" id="acessorioEstilo">
                                    <option value="">Selecione o estilo...</option>
                                </select>
                            </div>

                            <div class="col-lg-4 col-md-6">
                                <label for="acessorioNumeroCortes" class="form-label">
                                    Número de Cortes <span class="required-field">*</span>
                                </label>
                                <input type="number" class="form-control form-control-sm" id="acessorioNumeroCortes" min="0">
                            </div>

                            <div class="col-lg-4 col-md-6">
                                <label for="acessorioEstrutura" class="form-label">
                                    Estrutura <span class="required-field">*</span>
                                </label>
                                <select class="form-select form-select-sm select2" id="acessorioEstrutura">
                                    <option value="">Selecione a estrutura...</option>
                                </select>
                            </div>

                            <div class="col-lg-4 col-md-6">
                                <label for="acessorioLargura" class="form-label">Largura (mm) <span class="required-field">*</span></label>
                                <input type="number" class="form-control form-control-sm" id="acessorioLargura" min="0" step="0.01">
                            </div>

                            <div class="col-lg-4 col-md-6">
                                <label for="acessorioComprimento" class="form-label">Comprimento (mm) <span class="required-field">*</span></label>
                                <input type="number" class="form-control form-control-sm" id="acessorioComprimento" min="0" step="0.01">
                            </div>
                        </div>
                    </form>
                </div>
                <div class="modal-footer">
                    <button type="button" class="btn btn-sm btn-outline-secondary" data-bs-dismiss="modal">
                        Cancelar
                    </button>
                    <button type="button" class="btn btn-sm btn-primary" id="btnSalvarAcessorio">
                        <i class="bi bi-plus-lg me-1"></i>Adicionar
                    </button>
                </div>
            </div>
        </div>
    </div>

    <!-- Modal de Tipos de Produto Primário -->
    <div class="modal fade" id="modalTiposProduto" tabindex="-1" aria-labelledby="modalTiposProdutoLabel" aria-hidden="true">
        <div class="modal-dialog modal-lg modal-dialog-centered">
            <div class="modal-content">
                <div class="modal-header">
                    <h5 class="modal-title" id="modalTiposProdutoLabel">
                        <i class="bi bi-info-circle me-2"></i>Tipos de Produto Primário
                    </h5>
                    <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Fechar"></button>
                </div>
                <div class="modal-body d-flex justify-content-center">
                    <img src="http://br-pir1-m012/corporativo/ddpV1/imagens/Tipos.png" alt="Tipos de Produto Primário" class="img-fluid">
                </div>
                <div class="modal-footer">
                    <button type="button" class="btn btn-sm btn-secondary" data-bs-dismiss="modal">
                        Fechar
                    </button>
                </div>
            </div>
        </div>
    </div>
</body>
</html>