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

' Variáveis globais
Set usuarioLogado = CreateObject("Scripting.Dictionary")
usuarioLogado.Add "id", Session("CODIGO")
usuarioLogado.Add "nome", Session("NOME")
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
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/5.15.3/css/all.min.css" />

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

        /* Tabelas */
        #gradeProdutos th,
        #gradeProdutos td {
            white-space: nowrap;
            vertical-align: top;
        }

        thead {
            vertical-align: middle !important;
            position: sticky; 
            inset-block-start: 0; 
            background: #f5f5f5;
            z-index: 999;
        }

        .table-wrapper {
            overflow: auto;
            max-height: 55vh;
        }

        #gradeProdutos th:last-child,
        #gradeProdutos td:last-child {
            position: sticky;
            right: 0;
            z-index: 5;
            white-space: nowrap;
        }
        
        #gradeProdutos td:last-child {
            background: linear-gradient(to right, rgba(0,0,0,0.1) 0%, transparent 1px),#fff;
            background-origin: border-box;
            background-clip: padding-box;
        }
        
        #gradeProdutos thead th:last-child {
            background: linear-gradient(to right, rgba(0,0,0,0.1) 0%, transparent 1px), #f5f5f5;
            background-origin: border-box;
            background-clip: padding-box;
            z-index: 10;
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

        .truncate-text {
            max-width: 250px;
            white-space: nowrap;
            overflow: hidden;
            text-overflow: ellipsis;
            cursor: help;
        }

        .mensagem-detalhe-cell {
            min-width: 260px;
            vertical-align: top;
        }

        .mensagem-truncate-text {
            display: -webkit-box;
            -webkit-line-clamp: 2;
            -webkit-box-orient: vertical;
            max-width: 420px;
            overflow: hidden;
            text-overflow: ellipsis;
            white-space: normal;
            word-break: break-word;
        }

        .mensagem-detalhe-link {
            cursor: pointer;
            transition: color .15s ease;
        }

        .mensagem-detalhe-link:hover,
        .mensagem-detalhe-link:focus {
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
            display: none;
            padding: 30px;
            animation: fadeIn .25s ease;
            height: 80vh;
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
                font-size: 13px;
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
        .odd-box {
            border-radius: 6px;
            height: 25px;
            padding: 6px 10px;
            display: flex;
            align-items: center;
            gap: 5px;
            width: fit-content;
        }
        .odd-box.up {
            background-color: #d1fae5;
            color: #065f46;
        }
        .odd-box.down {
            background-color: #fee2e2;
            color: #991b1b;
        }
    </style>

    <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.0.2/dist/js/bootstrap.bundle.min.js" integrity="sha384-MrcW6ZMFYlzcLA8Nl+NtUVF0sA7MsXsP1UyJoMp4YLEuNSfAP+JcXn/tWtIaxVXM" crossorigin="anonymous"></script>
    <script src="https://cdn.jsdelivr.net/npm/select2@4.1.0-rc.0/dist/js/select2.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/jquery.mask/1.14.15/jquery.mask.min.js"></script>
    <script src="https://unpkg.com/just-validate@latest/dist/just-validate.production.min.js"></script>

    <script src="/corporativo/funcoes/funcoesJAVASCRIPT.js?v=<%=Now()%>"></script>
    <script src="/corporativo/components/confirmation.js?v=<%=Now()%>"></script>
    <script src="/corporativo/components/toast.js?v=<%=Now()%>"></script>
    <script src="/corporativo/components/dropdownMenu.js?v=<%=Now()%>"></script>

    <script>
        let numeroOrcamento = 0;
        let abaAtual = 1;
        let totalAbas = 4;
        let abasValidadas = new Set();
        let validacoes = {};
        let acao = false;
        let statusAtual = ''
        let modoVisualizacao = false;

        // Modal de Item
        let listaProdutos = [];
        let produtoEditandoIndex = null;
        let produtoVisualizando = false;
        var produtosSelecionados = new Map(); 

        var resultadosCalculo = {
            sucessoIds: [],
            erroIds: []
        };

        // Modal de Mensagens
        let listaMensagens = [];

        // Auxiliar para comparação de preços
        let listaUltimoPedidoFT = []

        // Busca de Cliente
        let clienteSelecionado = null;
        let resultadosCliente = [];

        // Busca de Vendedor
        let vendedorSelecionado = null;
        let resultadosVendedor = [];

        const eStatus = Object.freeze({
            APROVADO_TOTAL: 'Aprovado Total',
            CANCELADO: 'Cancelado',
            EM_ANALISE: 'Em Analise',
            EM_APROVACAO: 'Em Aprovacao',
            EM_DIGITACAO: 'Em Digitacao',
        });

        const eTipoItem = Object.freeze({
            CAIXA: 'CX',
            CHAPA: 'CH',
            CONJUNTO: 'CJ',
            ACESSSORIO: 'AC'
        })

        const eTipoCaixa = Object.freeze({
            MALETA_FUNDO: '1',
            MALETA_CAIXA_NORMAL: '2',
            MALETA_ABAS_TOTAL: '4',
            CORTE_VINCO: '5',
            CHAPAS: '6',
        })

        const eTipoPreco = Object.freeze({
            PRECO_UNITARIO: '1',
            PRECO_KG: '2',
            PRECO_M2: '3'
        })

        function exibirLoading(exibir) {
            carregando = exibir;
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
                dropdownCssClass: "z-index-9999" }
            );
        }

        function inicializarMascaras() {
            $('#CEP').mask('00000-000');
            $('#telefone').mask('(00) 00000-0000');
          
            $('#CNPJ').mask('AA.AAA.AAA/AAAA-00', {
                translation: { 'A': { pattern: /[A-Za-z0-9]/ } },
                onKeyPress: function(value, event, field) {
                    field.val(value.toUpperCase());
                }
            });
            
            $('#CNPJ').on('input', function() {
                this.value = this.value.toUpperCase();
            });
        }

        function carregarMunicipios(codPlanta, codMunicipio, UF) {
            const $select = $('#municipioId');
            $select.empty().append('<option value="">Selecione um município</option>');

            if(!codPlanta) {
                return;
            }

            const params = {
                acao: 'listarMunicipios',
                codPlanta: codPlanta,
                codMunicipio: codMunicipio,
                UF: UF
            };

            return $.ajax({
                url: 'OrcamentoRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: params,
                dataType: 'json',
                success: function(response) {
                    if (!response.sucesso) {
                        console.error("Erro ao carregar municípios: " + response.mensagem);
                        return;
                    }

                    response.dados.forEach(function(item) {
                        const selecionado = item.codMunicipio === codMunicipio;
                        $select.append($('<option>', { value: item.municipioId, text: item.nome, selected: selecionado }));
                    });
                },
                error: function(xhr, status, error) {
                    console.error('Erro ao carregar municípios:', error);
                }
            });
        }

        function carregarCondicoesPagamento(codPlanta) {
            if(!codPlanta) {
                return;
            }

            const $select = $('#condicaoPagamentoId');
            $select.empty().append('<option value="">Selecione uma condição de pagamento</option>');

            const params = {
                acao: 'listarCondicoesPagamento',
                codPlanta: codPlanta,
            };

            return $.ajax({
                url: 'OrcamentoRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: params,
                dataType: 'json',
                success: function(response) {
                    if (!response.sucesso) {
                        console.error("Erro ao carregar condições de pagamento: " + response.mensagem);
                        return;
                    }

                    response.dados.forEach(function(item) {
                        $select.append($('<option>', { value: item.codigo, text: item.descricao }));
                    });
                },
                error: function(xhr, status, error) {
                    console.error('Erro ao carregar condições de pagamento:', error);
                }
            });
        }

        function carregarTiposCaminhao(codPlanta) {
            if(!codPlanta) {
                return;
            }

            const $select = $('#tipoCaminhao');
            $select.empty().append('<option value="">Selecione um tipo de caminhão</option>');

            const params = {
                acao: 'listarTiposCaminhao',
                codPlanta: codPlanta,
            };

            return $.ajax({
                url: 'OrcamentoRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: params,
                dataType: 'json',
                success: function(response) {
                    if (!response.sucesso) {
                        console.error("Erro ao carregar condições de pagamento: " + response.mensagem);
                        return;
                    }

                    response.dados.forEach(function(item) {
                        $select.append($('<option>', { value: item.codigo, text: item.descricao }));
                    });
                },
                error: function(xhr, status, error) {
                    console.error('Erro ao carregar condições de pagamento:', error);
                }
            });
        }
    
        function carregarEstruturas(codPlanta) {
            if(!codPlanta) {
                return;
            }

            const $select = $('#produtoEstrutura');
            $select.empty().append('<option value="">Selecione uma estrutura</option>');

            const params = {
                acao: 'listarEstruturas',
                codPlanta: codPlanta,
            };

            return $.ajax({
                url: 'OrcamentoRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: params,
                dataType: 'json',
                success: function(response) {
                    if (!response.sucesso) {
                        console.error("Erro ao carregar estruturas: " + response.mensagem);
                        return;
                    }

                    response.dados.forEach(function(item) {
                        $select.append($('<option>', { value: item.codigo, text: item.codigo, 'gramatura': item.gramatura }));
                    });
                },
                error: function(xhr, status, error) {
                    console.error('Erro ao carregar estruturas:', error);
                }
            });
        }

        function carregarParametrosProduto(codPlanta) {
            if(!codPlanta) return

            const params = {
                acao: 'listarParametrosProduto',
                codPlanta: codPlanta,
            };

            return $.ajax({
                url: 'OrcamentoRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: params,
                dataType: 'json',
                success: function(response) {
                    if (!response.sucesso) {
                        console.error("Erro ao buscar parâmetros de tipo de produto: " + response.mensagem);
                        return;
                    }

                    $('#hidAreaMinimaMaleta').val(response.dados.areaMinimaMaleta);
                    $('#hidAreaMinimaCorteVinco').val(response.dados.areaMinimaCorteVinco);
                    $('#hidAreaMinimaChapa').val(response.dados.areaMinimaChapa);
                },
                error: function(xhr, status, error) {
                    console.error('Erro ao buscar parâmetros de tipo de produto:', error);
                }
            });
        }

        function calcularQuantidadeMinimaProduto() {
            const ehProdutoNovo = $('#flagProdutoNovo').is(':checked');
            const areaBlank = $('#hidProdutoAreaBlank').val()
            const areaRefile = $('#produtoAreaRefile').val()
            const valorArea = ehProdutoNovo ? parseFloat(areaRefile): parseFloat(areaBlank)

            const tipoProduto = $('#hidProdutoTipo').val() 
            const tipoCaixa = $('#produtoTipoCaixa').val()

            const quantidadeMinimaAnterior = $('#produtoQuantidade').val() || 0
            const areaMinimaMaleta = $('#hidAreaMinimaMaleta').val() || 0
            const areaMinimaCorteVinco = $('#hidAreaMinimaCorteVinco').val() || 0
            const areaMinimaChapa = $('#hidAreaMinimaChapa').val() || 0
            
            let quantidadeMinima = 0

            // Para produto existente usamos tipo produto para calcular
            if (!ehProdutoNovo && valorArea > 0) {
                if (tipoProduto == 'CHAPA') {
                    quantidadeMinima = Math.ceil(areaMinimaChapa / valorArea)
                } else if (tipoProduto == 'CORTE VINCO') {
                    quantidadeMinima = Math.ceil(areaMinimaCorteVinco / valorArea)
                } else if (tipoProduto == 'MALETA') {
                    quantidadeMinima = Math.ceil(areaMinimaMaleta / valorArea)
                } 
            }

            // Produto novo utilizamos campo de tipo de caixa
            if (ehProdutoNovo && valorArea > 0) {
                if (tipoCaixa == eTipoCaixa.MALETA_FUNDO) {
                    quantidadeMinima = Math.ceil(areaMinimaMaleta / valorArea)
                } else if (tipoCaixa == eTipoCaixa.MALETA_CAIXA_NORMAL) {
                    quantidadeMinima = Math.ceil(areaMinimaMaleta / valorArea)
                } else if (tipoCaixa == eTipoCaixa.MALETA_ABAS_TOTAL) {
                    quantidadeMinima = Math.ceil(areaMinimaMaleta / valorArea)
                } else if (tipoCaixa == eTipoCaixa.CORTE_VINCO) {
                    quantidadeMinima = Math.ceil(areaMinimaCorteVinco / valorArea)
                } else if (tipoCaixa == eTipoCaixa.CHAPAS) {
                    quantidadeMinima = Math.ceil(areaMinimaChapa / valorArea)
                }
            }

            $('#hidProdutoQuantidadeMinima').val(quantidadeMinima)
            
            if (quantidadeMinimaAnterior < quantidadeMinima) {
                $('#produtoQuantidade').val(quantidadeMinima)
                Toast.info(`Quantidade mínima alterada de ${quantidadeMinimaAnterior} para  ${quantidadeMinima}`, 'Atenção!', 5000)
            }
        }

        /**
         * Auxiliar para calculo de área de refile
         * Para tipo de caixa corte e vinco utilizamos comprimento da chapa, largura da chapa e arranjo
         * Outros tipos de caixa utilizamos comprimento da chapa e largura da chapa
        */
        function calcularAreaRefile() {
            $('#produtoAreaRefile').val('0');

            const tipoCaixa = $('#produtoTipoCaixa').val();
            const comprimentoChapa = $('#produtoComprimentoChapa').val();
            const larguraChapa = $('#produtoLarguraChapa').val();
            const arranjo = $('#produtoArranjo').val();
            const pecasPorConjunto = $('#produtoPecasPorConjunto').val();

            if(!comprimentoChapa || !larguraChapa || (tipoCaixa == eTipoCaixa.CORTE_VINCO && !arranjo)) return

            const resultado = tipoCaixa == eTipoCaixa.CORTE_VINCO ?
                ((comprimentoChapa / 1000) * (larguraChapa / 1000) * pecasPorConjunto) / arranjo :
                (comprimentoChapa / 1000) * (larguraChapa / 1000) * pecasPorConjunto;

            $('#produtoAreaRefile').val(trucarValor(resultado, 3)).trigger('input');
        }

        function calcularPesoRefile() {
            $('#produtoPesoRefile').val('0');
            
            const gramatura = $('#produtoGramatura').val();
            const area = $('#produtoAreaRefile').val();

            if(!gramatura || !area) return;

            const resultado = gramatura / 1000 * area;
            $('#produtoPesoRefile').val(trucarValor(resultado, 3));
        }


        function obterCliente() {
            const codCliente = $('#codCliente').val().trim();
            const codLoja = $('#codLoja').val();
            const codPlanta = $('#codPlanta').val();

            if (!codPlanta) {
                Toast.warning('Selecione uma planta para buscar por cliente');
                return
            }

            if (!codCliente) {
                limparCamposCliente();
                return
            }

            const params = {
                acao: 'listarClientes',
                codPlanta: codPlanta,
                codCliente: codCliente,
                codLoja: codLoja
            };

            exibirLoading(true);

            $.ajax({
                url: 'OrcamentoRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: params,
                dataType: 'json',
                success: function(response) {
                    if (!response.sucesso) {
                        console.error("Erro ao carregar obter cliente: " + response.mensagem);
                        limparCamposCliente();
                        Toast.error(response.mensagem || 'Erro ao buscar cliente.');
                        return;
                    }

                    const cliente = (response.dados || [])[0];
                    if (!cliente) {
                        limparCamposCliente();
                        Toast.info('Nenhum cliente encontrado para o código informado.');
                        return;
                    }

                    preencherDadosCliente(cliente)
                },
                error: function(xhr, status, error) {
                    console.error('Erro ao carregar obter cliente:', error);
                    limparCamposCliente();
                    Toast.error('Erro ao buscar cliente.');
                },
                complete: function() {
                    exibirLoading(false);
                }
            });
        }

        function obterVendedor() {
            const codVendedor = $('#codVendedor').val().trim();
            const codPlanta = $('#codPlanta').val();

            if (!codPlanta) {
                Toast.warning('Selecione uma planta para buscar por vendedor');
                return
            }

            if (!codVendedor) {
                limparCamposVendedor();
                return
            }

            const params = {
                acao: 'listarVendedores',
                codPlanta: codPlanta,
                busca: codVendedor
            };

            exibirLoading(true);

            $.ajax({
                url: 'OrcamentoRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: params,
                dataType: 'json',
                success: function(response) {
                    if (!response.sucesso) {
                        console.error("Erro ao carregar obter vendedor: " + response.mensagem);
                        limparCamposVendedor();
                        Toast.error(response.mensagem || 'Erro ao buscar vendedor.');
                        return;
                    }

                    const vendedor = (response.dados || [])[0];
                    if (!vendedor) {
                        Toast.info('Nenhum vendedor encontrado para o código informado.');
                        return;
                    }

                    preencherDadosVendedor(vendedor)
                },
                error: function(xhr, status, error) {
                    console.error('Erro ao carregar obter vendedor:', error);
                    limparCamposVendedor();
                    Toast.error('Erro ao buscar vendedor.');
                },
                complete: function() {
                    exibirLoading(false);
                }
            });
        }

        function obterProduto() {
            const codCliente = $('#codCliente').val();
            const codLoja = $('#codLoja').val();
            const codPlanta = $('#codPlanta').val();

            const codProduto = $('#produtoCod').val().trim();
            const tipoItem = $('#produtoTipoItem').val();

            if (!codPlanta) {
                Toast.warning('Selecione uma planta para buscar por produto');
                return
            }

            if (!codCliente) {
                Toast.warning('Informe o código do cliente antes de buscar por produto');
                return
            }

            if (!codProduto) {
                return
            }

            const params = {
                acao: 'listarProdutos',
                tipoItem: tipoItem,
                codPlanta: codPlanta,
                codCliente: codCliente,
                codLoja: codLoja,
                codProduto: codProduto
            };

            exibirLoading(true);

            $.ajax({
                url: 'OrcamentoRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: params,
                dataType: 'json',
                success: function(response) {
                    if (!response.sucesso) {
                        console.error("Erro ao carregar obter produto: " + response.mensagem);
                        Toast.error(response.mensagem || 'Erro ao buscar produto.');
                        return;
                    }

                    const produto = (response.dados || [])[0];
                    if (!produto) {
                        Toast.info('Nenhum produto encontrado para o código informado.');
                        return;
                    }

                    preencherDadosProduto(produto);
                },
                error: function(xhr, status, error) {
                    console.error('Erro ao carregar obter produto:', error);
                    Toast.error('Erro ao buscar produto.');
                },
                complete: function() {
                    exibirLoading(false);
                }
            });
        }

        function campoPlantaAlterado() {
            const codPlanta = $('#codPlanta').val();

            carregarCondicoesPagamento(codPlanta);
            carregarTiposCaminhao(codPlanta);
            carregarEstruturas(codPlanta);
            carregarParametrosProduto(codPlanta);
            limparCamposCliente();
        }

        function campoUFAlterado() {
            const codPlanta = $('#codPlanta').val();
            const codMunicipio = $('#hidCodMunicipio').val();
            const UF = $('#UF').val();

            carregarMunicipios(codPlanta, codMunicipio, UF);
        }

        async function campoTemICMSAlterado() {
            const temICMS = $('#temICMS').val();
            const aliquotaDesabilitada = temICMS == 0;
            let valorAliquota = 0;

            if(!aliquotaDesabilitada) {
                const codCliente =  $('#codCliente').val()
                const codPlanta = $('#codPlanta').val()
                const UF = $('#UF').val()
                resposta = await ObterICMSCliente(codCliente, codPlanta, UF);

                if(resposta.sucesso) {
                    valorAliquota = resposta.valor
                }
            }
            
            const aliquotaICMS = aliquotaDesabilitada ? 0 : valorAliquota;
            $('#aliquotaICMS').val(aliquotaICMS);
        }

        function ObterICMSCliente(codCliente, codPlanta, UF) {
            exibirLoading(true);

            return new Promise((resolve, reject) => {
                $.ajax({
                    url: 'OrcamentoRepositorio.asp?v=<%=Now()%>',
                    type: 'GET',
                    data: {
                        acao: 'ObterICMSCliente',
                        codCliente: codCliente, 
                        codPlanta: codPlanta,
                        UF: UF,
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
        
        function campoProdutoEstruturaAlterado() {
            const gramatura = $("#produtoEstrutura").find(":selected").attr("gramatura");
            $('#produtoGramatura').val(gramatura).trigger('input');
        }

        // Ao mudar valor do tipo de produto, precisamos limpar campos posteriores 
        // e filtrar as opções do combo 'Tipo de Caixa'
        function campoProdutoTipoAlterado() {
            const ehProdutoNovo = $('#flagProdutoNovo').is(':checked');
            const $produtoTipoItem = $('#produtoTipoItem');
            const $tipoCaixa = $('#produtoTipoCaixa');
            
            // Limpa campos de produto
            $tipoCaixa.empty();
            $('#hidProdutoIdOrcamento').val('');
            $('#produtoCod').val(ehProdutoNovo ? '1' : '');
            $('#produtoEstrutura').val('').trigger('change.select2');
            $('#produtoTipoPreco').val('1').trigger('change.select2');
            $('#produtoReferencia').val('');
            $('#produtoQuantidade').val('0');
            $('#produtoValorUnitario').val('0');
            $('#produtoQuantidadeCores').val('0');
            $('#produtoGramatura').val('0');
            $('#produtoComprimentoInterno').val('0');
            $('#produtoLarguraInterna').val('0');
            $('#produtoAlturaInterna').val('0');
            $('#produtoComprimentoChapa').val('0');
            $('#produtoLarguraChapa').val('0');
            $('#produtoArranjo').val('1');
            $('#produtoPecasPorConjunto').val('1');
            $('#produtoAreaRefile').val('0');
            $('#produtoPesoRefile').val('0');
            $('#hidProdutoAreaBlank').val('0');
            $('#hidProdutoPesoBlank').val('0');
            $('#produtoAliquotaIPI').val('0');
            $('#hidProdutoAliquotaIPI').val('0');
            $('#produtoPedidoCompra').val('');
            $('#produtoItemPedido').val('');
            $('#produtoTemIPI').val('0').trigger('change.select2');
            $('#produtoObservacao').val('');
            $('#hidProdutoQuantidadeMinima').val('0');
            $('#hidProdutoPrecoNet').val('0');
            $('#hidPrecoUnitario').val('0');
            $('#hidPrecoUnitarioNet').val('0');
            $('#hidProdutoPrecoKGRefile').val('0');
            $('#hidProdutoPrecoKGNetRefile').val('0');
            $('#hidProdutoPrecoKGNet').val('0');
            $('#hidProdutoVolumeM2').val('0');

            const tiposCaixaMap = {
                'CX': [
                    { value: '1', text: 'Maleta - Fundo (0200)' }, 
                    { value: '2', text: 'Maleta - Caixa Normal' }, 
                    { value: '4', text: 'Maleta - Abas Total (0203)' }, 
                    { value: '5', text: 'Corte e Vinco' }, 
                ],
                'CJ': [
                    { value: '1', text: 'Maleta - Fundo (0200)' }, 
                    { value: '2', text: 'Maleta - Caixa Normal' }, 
                    { value: '4', text: 'Maleta - Abas Total (0203)' }, 
                    { value: '5', text: 'Corte e Vinco' }, 
                ],
                'CH': [ 
                    { value: '6', text: 'Chapas' } 
                ],
                'AC': [
                    { value: '1', text: 'Maleta - Fundo (0200)' }, 
                    { value: '2', text: 'Maleta - Caixa Normal' }, 
                    { value: '4', text: 'Maleta - Abas Total (0203)' }, 
                    { value: '5', text: 'Corte e Vinco' }, 
                ]
            }

            const itens = tiposCaixaMap[$produtoTipoItem.val()] || [];
            const options = itens.map((item, index) => `<option ${index === 0 ? 'selected' : ''} value="${item.value}">${item.text}</option>`).join('');
            $tipoCaixa.append(options).trigger('change.select2');

            aplicarEstadoCamposProduto();
            atualizarLabelValorUnitario();
            atualizarMarcacoesObrigatorias();
        }

        function campoProdutoTipoCaixaAlterado() {
            calcularQuantidadeMinimaProduto();
            atualizarMarcacoesObrigatorias();
        }

        function atualizarLabelValorUnitario() {
            const tipoPreco = $('#produtoTipoPreco').val();
            const labels = {
                '1': 'Valor Unitário',
                '2': 'Valor KG',
                '3': 'Valor M2'
            };
            $('#produtoValorUnitarioLabel').html(`${labels[tipoPreco] || 'Valor Unitário'} <span class="required-field">*</span>`);
        }

        function campoProdutoTemIPIAlterado() {
            let valorAliquota = $('#produtoTemIPI').val() == '1' ? $('#hidProdutoAliquotaIPI').val() : 0
            $('#produtoAliquotaIPI').val(valorAliquota);
            atualizarMarcacoesObrigatorias();
        }

        function abrirModalCliente() {
            const codPlanta = $('#codPlanta').val();

            if (!codPlanta) { 
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
            const codPlanta = $('#codPlanta').val();

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
                url: 'OrcamentoRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: {
                    acao: 'listarClientes',
                    codPlanta: codPlanta,
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
                        <td>${cliente.codCliente}</td>
                        <td>${cliente.codLoja}</td>
                        <td>${formatarCNPJ(cliente.CNPJ)}</td>
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

        function abrirModalVendedor() {
            const codPlanta = $('#codPlanta').val();

            if (!codPlanta) { 
                Toast.warning('Informe o código da planta antes de buscar por vendedores.'); 
                return; 
            }

            vendedorSelecionado = null;
            resultadosVendedor = [];
            $('#vendedorBusca').val('');
            $('#vendedorResultados').html(`
                <div class="text-center text-muted">
                    <i class="bi bi-search" style="font-size: 1.5rem;"></i>
                    <p class="mt-2 mb-0">Digite um termo para buscar vendedores existentes.</p>
                </div>
            `);
            $('#btnConfirmarVendedor').prop('disabled', true);

            const modal = new bootstrap.Modal(document.getElementById('modalVendedor'));
            modal.show();

            setTimeout(function() { 
                $('#vendedorBusca').focus(); 
            }, 500);
        }

        function pesquisarVendedor() {
            const busca = $('#vendedorBusca').val().trim();
            const codPlanta = $('#codPlanta').val();

            if (!busca) {
                $('#vendedorResultados').html(`
                    <div class="text-center text-muted">
                        <i class="bi bi-search" style="font-size: 1.5rem;"></i>
                        <p class="mt-2 mb-0">Digite um termo para buscar vendedores existentes.</p>
                    </div>
                `);
                return;
            }

            vendedorSelecionado = null;
            $('#btnConfirmarVendedor').prop('disabled', true);

            $('#vendedorResultados').html(`
                <div class="text-center">
                    <div class="spinner-border spinner-border-sm text-primary" role="status">
                        <span class="visually-hidden">Buscando...</span>
                    </div>
                    <p class="mt-2 mb-0 text-muted">Buscando vendedores...</p>
                </div>
            `);

            $.ajax({
                url: 'OrcamentoRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: {
                    acao: 'listarVendedores',
                    codPlanta: codPlanta,
                    busca: busca
                },
                dataType: 'json',
                success: function(response) {
                    if (!response.sucesso) {
                        $('#vendedorResultados').html(`
                            <div class="alert text-center alert-danger" role="alert">
                                <i class="bi bi-exclamation-triangle" style="font-size: 1.5rem;"></i>
                                <p class="mt-2 mb-0"> ${response.mensagem || 'Erro ao buscar vendedores.'} </p>
                            </div>
                        `);
                        return;
                    }

                    resultadosVendedor = response.dados || [];

                    if (resultadosVendedor.length === 0) {
                        $('#vendedorResultados').html(`
                            <div class="alert text-center alert-info" role="alert">
                                <i class="bi bi-info-circle" style="font-size: 1.5rem;"></i>
                                <p class="mt-2 mb-0"> Nenhum vendedor encontrado com os termos pesquisados. </p>
                            </div>
                        `);
                        return;
                    }

                    renderizarResultadosVendedor();
                },
                error: function(xhr, status, error) {
                    console.error('Erro ao buscar vendedores:', error);
                    $('#vendedorResultados').html(`
                        <div class="alert text-center alert-info" role="alert">
                            <i class="bi bi-info-circle" style="font-size: 1.5rem;"></i>
                            <p class="mt-2 mb-0"> Nenhum vendedor encontrado com os termos pesquisados.</p>
                        </div>
                    `);
                }
            });
        }

        function renderizarResultadosVendedor() {
            let html = `
                <div class="table-responsive">
                    <table class="table table-sm table-hover mb-0">
                        <thead class="thead-light">
                            <tr>
                                <th style="width: 50px;"></th>
                                <th>Código</th>
                                <th>Nome</th>
                                <th>CNPJ</th>
                            </tr>
                        </thead>
                        <tbody>
            `;

            resultadosVendedor.forEach(function(vendedor, index) {
                html += `
                    <tr data-index="${index}">
                        <td class="text-center">
                            <button onclick="selecionarVendedor(${index})"
                                class="btn-selecionar" 
                                title="Selecionar este vendedor">
                                <i class="bi bi-circle"></i> 
                            </button>
                        </td>
                        <td>${vendedor.codigo}</td>
                        <td>${vendedor.nome}</td>
                        <td>${formatarCNPJ(vendedor.cnpj)}</td>
                    </tr>
                `;
            });

            html += `
                        </tbody>
                    </table>
                </div>
            `;

            $('#vendedorResultados').html(html);
        }

        function selecionarVendedor(index) {
            vendedorSelecionado = resultadosVendedor[index];
            
            $('#vendedorResultados .btn-selecionar')
                .removeClass('selecionado')
                .html('<i class="bi bi-circle"></i>')
                .attr('title', 'Selecionar este item');
            
            $('#vendedorResultados tr[data-index="' + index + '"] .btn-selecionar')
                .addClass('selecionado')
                .html('<i class="bi bi-check-circle-fill"></i>')
                .attr('title', 'Item Selecionado');
            
            $('#btnConfirmarVendedor').prop('disabled', false);
        }

        function confirmarVendedor() {
            if (!vendedorSelecionado) {
                Toast.warning('Selecione um vendedor');
                return;
            }

            const modal = bootstrap.Modal.getInstance(document.getElementById('modalVendedor'));
            modal.hide();

            preencherDadosVendedor(vendedorSelecionado);
        }

        function preencherDadosVendedor(dados) {
            $('#codVendedor').val(dados.codigo);
            $('#nomeVendedor').val(dados.nome);
        }

        function limparCamposVendedor() {
            $('#nomeVendedor').val('');
        }

        function preencherDadosCliente(dados, atualizarMunicipios = true, aplicarEstado = true) {
            const codMunicipio = dados.codMunicipio || $('#hidCodMunicipio').val() || '';

            $('#nomeCliente').val(dados.nome || dados.nomeCliente || '');
            $('#codCliente').val(dados.codCliente || '');
            $('#codLoja').val(dados.codLoja || '');
            $('#CNPJ').val(dados.CNPJ || '').trigger('input');
            $('#codVendedor').val(dados.codVendedor || '');
            $('#nomeVendedor').val(dados.nomeVendedor || '');
            $('#email').val(dados.email || '');
            $('#endereco').val(dados.endereco || '');
            $('#numero').val(dados.numero || '');
            $('#complemento').val(dados.complemento || '');
            $('#UF').val(dados.UF || '').trigger('change.select2');
            $('#hidCodMunicipio').val(codMunicipio);
            $('#bairro').val(dados.bairro || '');
            $('#CEP').val(dados.CEP || '').trigger('input');
            $('#telefone').val(dados.telefone || '');
            $('#contato').val(dados.contato || '');

            if (dados.municipioId) {
                $('#municipioId').val(dados.municipioId).trigger('change.select2');
            }

            if(dados.codCliente == '1') {
                $('#flagClienteNovo').prop('checked', true);
            }

            if (atualizarMunicipios) {
                const codPlanta = $('#codPlanta').val();
                carregarMunicipios(codPlanta, codMunicipio, dados.UF);
            }

            if (aplicarEstado) {
                aplicarEstadoCamposCliente();
            }
        }

        /* Aba de Produtos */
        function abrirModalPesquisaProduto() {
            if ($('#flagClienteNovo').is(':checked')) {
                Toast.warning('Para clientes novos, somente é permitido cadastrar produto novo.');
                return;
            }

            const codPlanta = $('#codPlanta').val();
            const codCliente = $('#codCliente').val();

            if (!codPlanta) { 
                Toast.warning('Informe o código da planta antes de buscar por produtos.'); 
                return; 
            }

            if (!codCliente) {
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
            const tipoItem = $('#produtoTipoItem').val();
            const codPlanta = $('#codPlanta').val();
            const codCliente = $('#codCliente').val();
            const codLoja = $('#codLoja').val();

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
                url: 'OrcamentoRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: {
                    acao: 'listarProdutos',
                    tipoItem: tipoItem,
                    codPlanta: codPlanta,
                    codCliente: codCliente,
                    codLoja: codLoja,
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
                                <th>Descrição</th>
                                <th>Estrutura</th>
                                <th>Status</th>
                                <th>Cliente</th>
                            </tr>
                        </thead>
                        <tbody>
            `;

            resultadosProduto.forEach(function(produto, index) {
                const statusClasse = produto.status === 'ATIVO'  ? 'badge-green' : (produto.status === 'INATIVO' ? 'badge-red' : 'badge-yellow');

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
                        <td>${produto.descricao}</td>
                        <td>${produto.estrutura}</td>
                        <td><span class="badge ${statusClasse}">${produto.status || '-'}</span></td>
                        <td>${produto.codCliente}</td>
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
            $('#flagProdutoNovo').prop('checked', !!parseInt(dados.flagProdutoNovo || 0));
                        
            if(dados.tipoItem)  {
                $('#produtoTipoItem').val(dados.tipoItem).trigger('change.select2');
                campoProdutoTipoAlterado();
            }

            if(dados.tipoCaixa) $('#produtoTipoCaixa').val(dados.tipoCaixa).trigger('change.select2');
            if(dados.tipoPreco) $('#produtoTipoPreco').val(dados.tipoPreco).trigger('change.select2');

            $('#produtoCod').val(dados.codigo);
            $('#produtoEstrutura').val(dados.estrutura).trigger('change.select2');
            $('#produtoReferencia').val(dados.referencia);
            $('#produtoQuantidade').val(dados.quantidade);
            $('#produtoValorUnitario').val(dados.valorUnitario);
            $('#produtoQuantidadeCores').val(dados.quantidadeCores || 0);
            $('#produtoGramatura').val(dados.gramatura || 0);
            $('#produtoComprimentoInterno').val(dados.comprimentoInterno || 0);
            $('#produtoLarguraInterna').val(dados.larguraInterna || 0);
            $('#produtoAlturaInterna').val(dados.alturaInterna || 0);
            $('#produtoComprimentoChapa').val(dados.comprimentoChapa || 0);
            $('#produtoLarguraChapa').val(dados.larguraChapa || 0);
            $('#produtoArranjo').val(dados.arranjo || 1);
            $('#produtoPecasPorConjunto').val(dados.pecasPorConjunto || 1);
            $('#produtoAreaRefile').val(dados.areaRefile || 0);
            $('#produtoPesoRefile').val(dados.pesoRefile || 0);
            $('#hidProdutoAreaBlank').val(dados.areaBlank || 0);
            $('#hidProdutoPesoBlank').val(dados.pesoBlank || 0);
            $('#produtoAliquotaIPI').val(dados.aliquotaIPI || 0);
            $('#hidProdutoAliquotaIPI').val((dados.aliquotaIPIOriginal || dados.aliquotaIPI || 0));
            $('#produtoPedidoCompra').val(dados.pedidoCompra);
            $('#produtoItemPedido').val(dados.itemPedido);
            $('#produtoTemIPI').val(dados.aliquotaIPI > 0 || dados.temIPI != 0 ? 1 : 0).trigger('change.select2');
            $('#produtoObservacao').val(dados.observacao);
            
            // Campos ocultos
            $('#hidProdutoTipo').val(dados.tipo);
            $('#hidProdutoIdOrcamento').val(dados.idOrcamento);
            $('#hidProdutoStatusCalculo').val(dados.statusCalculo);
            $('#hidProdutoPrecoNet').val(dados.precoNet || 0);
            $('#hidPrecoUnitario').val(dados.precoUnitario);
            $('#hidPrecoUnitarioNet').val(dados.precoUnitarioNet);
            $('#hidProdutoPrecoKGRefile').val(dados.precoKGRefile || 0);
            $('#hidProdutoPrecoKGNetRefile').val(dados.precoKGNetRefile || 0);
            $('#hidProdutoPrecoKGNet').val(dados.precoKGNetBlank || 0);
            $('#hidProdutoVolumeM2').val(dados.volumeM2 || 0);

            atualizarLabelValorUnitario();
            atualizarMarcacoesObrigatorias();
            calcularQuantidadeMinimaProduto();
            aplicarEstadoCamposProduto();
        }
                
        function abrirModalProduto() {            
            produtoEditandoIndex = null;
            produtoVisualizando = false;
            $('#btnSalvarProduto').removeClass('d-none');
                 
            const ehClienteNovo = $('#flagClienteNovo').is(':checked');
            $('#flagProdutoNovo').prop('checked', ehClienteNovo).prop('disabled', ehClienteNovo);

            limparCamposProduto();
            $('#modalProdutoTitulo').text('Adicionar Produto');
            $('#btnSalvarProduto').html('<i class="bi bi-plus-lg me-1"></i>Adicionar');
            
            const modal = new bootstrap.Modal(document.getElementById('modalProduto'));
            modal.show();

            validacoes.modalProduto.refresh();
            atualizarMarcacoesObrigatorias();
        }

        function editarProduto(index) {
            const produto = listaProdutos[index];
            if (!produto) return;

            produtoEditandoIndex = index;
            produtoVisualizando = false;
            $('#btnSalvarProduto').removeClass('d-none');
            preencherDadosProduto(produto);

            $('#modalProdutoTitulo').text('Editar Produto');
            $('#btnSalvarProduto').html('<i class="bi bi-check-lg me-1"></i>Salvar');

            const modal = new bootstrap.Modal(document.getElementById('modalProduto'));
            modal.show();
        }


        function detalhesProduto(idOrcamento, index) {
            const produto = listaProdutos[index] || listaProdutos.find(item => String(item.idOrcamento) === String(idOrcamento));
            if (!produto) return;

            produtoEditandoIndex = null;
            produtoVisualizando = true;

            preencherDadosProduto(produto);
            $('#modalProdutoTitulo').text('Detalhes do Produto');
            aplicarModoVisualizacaoProduto(true);

            const modal = new bootstrap.Modal(document.getElementById('modalProduto'));
            modal.show();
        }

        function aplicarModoVisualizacaoProduto(visualizando) {
            produtoVisualizando = visualizando;
            $('#btnSalvarProduto').toggleClass('d-none', visualizando);

            if (!visualizando) return;

            $('#formularioProduto').find('input, select, textarea, button').prop('disabled', true);
            $('#formularioProduto select').trigger('change.select2');
        }

        function removerProduto(index) {
            ConfirmationModal.danger({
                message: 'Deseja realmente remover este produto?',
                type: 'danger',
                confirmText: 'Remover',
                cancelText: 'Cancelar',
                onConfirm: () => {
                    listaProdutos.splice(index, 1);
                    renderizarProdutos();
                    Toast.success('Produto removido com sucesso!');
                },
                onCancel: () => {}
            });
        }

        function salvarProduto() {   
            const ehProdutoNovo = $('#flagProdutoNovo').is(':checked') ? 1 : 0;
            const ehChapa = $('#produtoTipoCaixa').val() == eTipoCaixa.CHAPAS;

            const areaRefile = parseFloat($('#produtoAreaRefile').val()) || 0
            const pesoRefile = parseFloat($('#produtoPesoRefile').val()) || 0
            
            // Existem três situações que precisamos replicar as medidas refile nas medidas blank:
            // Se for produto novo, se for chapa ou se medidas blank estiverem zeradas
            const areaBlank = (ehProdutoNovo || ehChapa || parseFloat($('#hidProdutoAreaBlank').val()) == 0) ? 
                areaRefile : parseFloat($('#hidProdutoAreaBlank').val()) 
                
            const pesoBlank = (ehProdutoNovo || ehChapa || parseFloat($('#hidProdutoPesoBlank').val()) == 0) ?
                pesoRefile : parseFloat($('#hidProdutoPesoBlank').val()) 
            
            const tipoPreco = $('#produtoTipoPreco').val();
            const valorUnitario = $('#produtoValorUnitario').val();

            // Campo valorUnitario é um campo genérico que vale para preco_unitario, preco_kg
            // Se tipoPreco = preco_unitario então campo precoUnitario recebe valorUnitario
            // Se tipoPreco = preco_kg então campo precoUnitario recebe valorUnitario
            const precoUnitario = tipoPreco == eTipoPreco.PRECO_UNITARIO ? valorUnitario : $('#hidPrecoUnitario').val() 
            const precoKGRefile = tipoPreco == eTipoPreco.PRECO_KG ? valorUnitario : $('#hidProdutoPrecoKGRefile').val() 

            calcularQuantidadeMinimaProduto()
            
            const produto = {
                flagProdutoNovo: ehProdutoNovo,
                idOrcamento: $('#hidProdutoIdOrcamento').val(),
                statusCalculo: $('#hidProdutoStatusCalculo').val(),
                codigo: $('#produtoCod').val(),
                estrutura: $('#produtoEstrutura').val(),
                tipoItem: $('#produtoTipoItem').val(),
                tipoCaixa: $('#produtoTipoCaixa').val(),
                tipoPreco: tipoPreco,
                referencia: $('#produtoReferencia').val(),
                quantidade: $('#produtoQuantidade').val(),
                valorUnitario: valorUnitario,
                quantidadeCores: $('#produtoQuantidadeCores').val(),
                gramatura: $('#produtoGramatura').val(),
                comprimentoInterno: $('#produtoComprimentoInterno').val(),
                larguraInterna: $('#produtoLarguraInterna').val(),
                alturaInterna: $('#produtoAlturaInterna').val(),
                comprimentoChapa: $('#produtoComprimentoChapa').val(),
                larguraChapa: $('#produtoLarguraChapa').val(),
                arranjo: $('#produtoArranjo').val(),
                pecasPorConjunto: $('#produtoPecasPorConjunto').val(),
                areaRefile: areaRefile,
                pesoRefile: pesoRefile,
                areaBlank: areaBlank,
                pesoBlank: pesoBlank,
                aliquotaIPI: $('#produtoAliquotaIPI').val(),
                aliquotaIPIOriginal: $('#hidProdutoAliquotaIPI').val(),
                pedidoCompra: $('#produtoPedidoCompra').val(),
                itemPedido: $('#produtoItemPedido').val(),
                temIPI: $('#produtoTemIPI').val(),
                observacao: $('#produtoObservacao').val(),
                precoNet: $('#hidProdutoPrecoNet').val(),
                precoUnitario: precoUnitario,
                precoUnitarioNet: $('#hidPrecoUnitarioNet').val(),
                precoKGRefile: precoKGRefile,
                precoKGNetRefile: $('#hidProdutoPrecoKGNetRefile').val(),
                precoKGNetBlank: $('#hidProdutoPrecoKGNet').val(),
                volumeM2: $('#hidProdutoVolumeM2').val()
            };

            if (produtoEditandoIndex !== null) {
                listaProdutos[produtoEditandoIndex] = produto;
                Toast.success('Produto atualizado com sucesso!');
            } else {
                listaProdutos.push(produto);
                Toast.success('Produto adicionado com sucesso!');
            }

            renderizarProdutos();

            const modal = bootstrap.Modal.getInstance(document.getElementById('modalProduto'));
            modal.hide();
        }

        /**
         * Auxiliar para limpar campos de produto
         * Se não for produto novo, desabilita os campos que não podem ser editados
         */
        function aplicarEstadoCamposProduto() {
            if (produtoVisualizando) {
                aplicarModoVisualizacaoProduto(true);
                return;
            }

            const editandoProduto = produtoEditandoIndex !== null;
            const ehClienteNovo = $('#flagClienteNovo').is(':checked');
            const ehProdutoNovo = $('#flagProdutoNovo').is(':checked');
            const ehChapa = $('#produtoTipoCaixa').val() == eTipoCaixa.CHAPAS;

            const camposProduto = `#produtoTipoItem, #produtoTipoCaixa, #produtoCod, #produtoEstrutura,
                #produtoReferencia, #produtoQuantidade, #produtoTipoPreco, #produtoValorUnitario,
                #produtoQuantidadeCores, #produtoGramatura, #produtoComprimentoInterno,
                #produtoLarguraInterna, #produtoAlturaInterna, #produtoComprimentoChapa,
                #produtoLarguraChapa, #produtoArranjo, #produtoPecasPorConjunto, #produtoAreaRefile, #produtoPesoRefile,
                #produtoTemIPI, #produtoAliquotaIPI, #produtoPedidoCompra, #produtoItemPedido,
                #produtoObservacao, #btnBuscarProduto`;

            const camposProdutoNovoDesabilitados = `#btnBuscarProduto, #produtoCod, #produtoGramatura,
                #produtoAreaRefile, #produtoPesoRefile`;

            const camposProdutoExistenteNaoChapaDesabilitados = `#produtoEstrutura, #produtoReferencia,
                #produtoGramatura, #produtoComprimentoInterno, #produtoLarguraInterna,
                #produtoAlturaInterna, #produtoComprimentoChapa, #produtoLarguraChapa,
                #produtoArranjo, #produtoPecasPorConjunto, #produtoAreaRefile, #produtoPesoRefile, #produtoAliquotaIPI,
                #produtoQuantidadeCores`;
            
            const camposProdutoExistenteChapaDesabilitados = `#produtoGramatura, #produtoAreaRefile, #produtoPesoRefile`;

            $(camposProduto).prop('disabled', false);

            if (editandoProduto) {
                $('#flagProdutoNovo').prop('disabled', true);
            } else {
                $('#flagProdutoNovo').prop('disabled', ehClienteNovo);
            }
            
            if(ehProdutoNovo) {
                // Campos desabilitados para produto novo
                $(camposProdutoNovoDesabilitados).prop('disabled', true);
            } else if(ehChapa) {
                // Campos desabilitados para produto existente == chapa
                $(camposProdutoExistenteChapaDesabilitados).prop('disabled', true);
            } else {
                // Campos desabilitados para produto existente != chapa
                $(camposProdutoExistenteNaoChapaDesabilitados).prop('disabled', true);
            }

            $('#formularioProduto select').trigger('change.select2');
        }

        /**
         * Auxiliar para limpar campos de produto
         * Aplica os campos habilitados/desabilitados conforme o cenário atual:
         * produto novo, produto existente ou edição de produto.
         */
        function limparCamposProduto() {
            const ehClienteNovo = $('#flagClienteNovo').is(':checked');
            if (ehClienteNovo && !$('#flagProdutoNovo').is(':checked')) {
                $('#flagProdutoNovo').prop('checked', true);
            }

            const ehProdutoNovo = $('#flagProdutoNovo').is(':checked');
            
            $('#produtoCod').val(ehProdutoNovo ? '1' : '');
            $('#btnBuscarProduto').prop('disabled', ehProdutoNovo);
            $('#produtoEstrutura').val('').trigger('change.select2');
            $('#produtoTipoItem').val('CX').trigger('change.select2');
            
            campoProdutoTipoAlterado();

            $('#produtoTipoCaixa').val('1').trigger('change.select2');
            $('#produtoTipoPreco').val('1').trigger('change.select2');
            $('#produtoReferencia').val('');
            $('#produtoQuantidade').val('0');
            $('#produtoValorUnitario').val('0');
            $('#produtoQuantidadeCores').val('0');
            $('#produtoGramatura').val('0');
            $('#produtoComprimentoInterno').val('0');
            $('#produtoLarguraInterna').val('0');
            $('#produtoAlturaInterna').val('0');
            $('#produtoComprimentoChapa').val('0');
            $('#produtoLarguraChapa').val('0');
            $('#produtoArranjo').val('1');
            $('#produtoPecasPorConjunto').val('1');
            $('#produtoAreaRefile').val('0');
            $('#produtoPesoRefile').val('0');
            $('#hidProdutoAreaBlank').val('0');
            $('#hidProdutoPesoBlank').val('0');
            $('#produtoAliquotaIPI').val('0');
            $('#hidProdutoAliquotaIPI').val('0');
            $('#produtoPedidoCompra').val('')
            $('#produtoItemPedido').val('')
            $('#produtoTemIPI').val('0').trigger('change.select2')
            $('#produtoObservacao').val('')
            
            // Campos ocultos
            $('#hidProdutoQuantidadeMinima').val('0');
            $('#hidProdutoPrecoNet').val('0');
            $('#hidPrecoUnitario').val('0');
            $('#hidPrecoUnitarioNet').val('0');
            $('#hidProdutoPrecoKGRefile').val('0');
            $('#hidProdutoPrecoKGNetRefile').val('0');
            $('#hidProdutoPrecoKGNet').val('0');
            $('#hidProdutoVolumeM2').val('0');
            $('#hidProdutoStatusCalculo').val('Pendente');
            $('#hidProdutoIdOrcamento').val('');
            
            atualizarLabelValorUnitario();
            atualizarMarcacoesObrigatorias();
            aplicarEstadoCamposProduto();
        }

        
        function aplicarEstadoCamposCliente() {
            const ehClienteNovo = $('#flagClienteNovo').is(':checked');

            const camposBuscaCliente = '#btnBuscarCliente, #codCliente, #codLoja';
            const camposCadastroCliente = `#nomeCliente, #CNPJ, #btnBuscarVendedor, #codVendedor, #nomeVendedor,
                #email, #CEP, #endereco, #bairro, #UF, #municipioId, #numero`;

            $(camposBuscaCliente).prop('disabled', false);
            $(camposCadastroCliente).prop('disabled', false);

            if (ehClienteNovo) {
                $(camposBuscaCliente).prop('disabled', true);
            } else {
                $(camposCadastroCliente).prop('disabled', true);
            }

            $('#UF, #municipioId').trigger('change.select2');
        }

        /**
         * Auxiliar para limpar campos de cliente.
         * A limpeza de valores fica separada da regra de habilitar/desabilitar campos.
         */
        function limparCamposCliente() {
            const ehClienteNovo = $('#flagClienteNovo').is(':checked');

            $('#codCliente').val(ehClienteNovo ? '1' : '');
            $('#codLoja').val(ehClienteNovo ? '01' : '');
            $('#nomeCliente').val('');
            $('#CNPJ').val('').trigger('input');
            $('#codVendedor').val('');
            $('#nomeVendedor').val('');
            $('#email').val('');
            $('#CEP').val('').trigger('input');
            $('#endereco').val('');
            $('#bairro').val('');
            $('#UF').val('').trigger('change.select2');
            $('#municipioId').val('').trigger('change.select2');
            $('#hidCodMunicipio').val('');
            $('#numero').val('');
            $('#complemento').val('');
            $('#telefone').val('');
            $('#contato').val('');

            aplicarEstadoCamposCliente();
        }

        function renderizarProdutos() {
            const $container = $('#produtoListaContainer');
            const $contador = $('#contadorProdutos, #menuContadorProdutos');
            
            $contador.text(listaProdutos.length);

            if (listaProdutos.length === 0) {
                $container.html(`
                    <div class="text-center py-5">
                        <p class="text-muted mt-3 mb-0">Nenhum produto adicionado.</p>
                    </div>
                `);
                return;
            }

            let html = `
                <div id="avisoCalculoProdutos" class="alert alert-info d-flex align-items-start gap-2 py-2" role="alert">
                    <i class="bi bi-info-circle-fill mt-1"></i>
                    <div>
                        Alguns valores podem aparecer zerados no momento da inclusão. Após o envio do formulário, o sistema fará as validações necessárias e atualizará automaticamente esses valores.
                    </div>
                </div>
                <div class="table-wrapper">
                    <table id="gradeProdutos" class="table table-hover table-sm mb-0">
                        <thead class="table-light">
                            <tr>
                                <th style="width: 50px;">
                                    <input type="checkbox" class="form-check-input" id="checkboxSelectAll" onclick="toggleSelectAll()">
                                </th>
                                <th>#</th>
                                <th>Calculo</th>
                                <th>Produto</th>
                                <th>Quantidade</th>
                                <th>Referência</th>
                                <th>Tipo de Item</th>
                                <th>Peças p/ Conjunto</th>
                                <th>Comprimento Interno</th>
                                <th>Largura Interna</th>
                                <th>Altura Interna</th>
                                <th>Comprimento Chapa</th>
                                <th>Largura Chapa</th>
                                <th>Estrutura</th>
                                <th>Gramatura</th>
                                <th>Área Refile</th>
                                <th>Peso Refile</th>
                                <th>Preço Unitário</th>
                                <th>Preço Unitário NET</th>
                                <th>Preço KG Refile</th>
                                <th>Preço KG Net Refile</th>
                                <th>Preço KG Net</th>
                                <th>Qtd Cores</th>
                                <th>Volume M2</th>
                                <th>Margem</th>
                                <th class="text-center"></th>
                            </tr>
                        </thead>
                        <tbody>
            `;

            
            listaProdutos.forEach((produto, index) => {
                const produtoSelecionado = produtosSelecionados.has(String(produto.idOrcamento));
                const comparacaoPreco = obterComparacaoPrecoKGNet(produto);
                const badgeCalculo = obterBadgeStatusCalculo(produto.statusCalculo);

                html += `
                    <tr>
                        <td> 
                            <input type="checkbox" class="form-check-input checkbox-produto" 
                            data-id-produto="${produto.idOrcamento}"
                            onclick="toggleSelecao('${produto.idOrcamento}')"
                            ${(produtoSelecionado ? 'checked' : '')} 
                            >
                        </td>
                        <td>${produto.idOrcamento || '0'}</>
                        <td> <span class="badge ${badgeCalculo.classe}">${badgeCalculo.texto}</span> </td>
                        <td>${produto.codigo}</td> 
                        <td>${produto.quantidade || 0}</td> 
                        <td>
                            <div class="truncate-text" title="${produto.referencia || ''}">
                                ${produto.referencia || '-'}
                            </div>
                        </td>
                        <td>${produto.tipoItem}</td> 
                        <td>${produto.pecasPorConjunto || 1}</td> 
                        <td>${produto.comprimentoInterno || 0}</td> 
                        <td>${produto.larguraInterna || 0}</td> 
                        <td>${produto.alturaInterna || 0}</td> 
                        <td>${produto.comprimentoChapa || 0}</td> 
                        <td>${produto.larguraChapa || 0}</td> 
                        <td>${produto.estrutura || 0}</td> 
                        <td>${produto.gramatura || 0}</td> 

                        <td>${produto.areaRefile || 0}</td> 
                        <td>${produto.pesoRefile || 0}</td> 

                        <td>${produto.precoUnitario || 0}</td> 
                        <td>${produto.precoUnitarioNet || 0}</td> 

                        <td>${produto.precoKGRefile || 0}</td> 
                        <td>${produto.precoKGNetRefile || 0}</td> 

                        <td>
                            <div class="odd-box ${comparacaoPreco.classe}" data-bs-toggle="tooltip" data-bs-placement="top" title="${comparacaoPreco.textoTooltip}">
                                <div>${produto.precoKGNetBlank || 0}</div>
                                <i class="${comparacaoPreco.icone}"></i>
                            </div>
                        </td>
                        
                        <td>${produto.quantidadeCores || 0}</td>
                        <td>${produto.volumeM2 || 0}</td>
                        <td>${produto.margem || 0}</td>
                        ${obterAcoesProduto(produto, index)}
                    </tr>
                `;
            });

            html += `
                        </tbody>
                    </table>
                </div>
            `;

            $container.html(html);
            atualizarCheckboxSelectAll();
        }


        function compararPrecoKGNet(precoAtual, precoAnterior) {
            if (precoAnterior === null || isNaN(parseFloat(precoAnterior))) {
                return { classe: 'stable', icone: '' };
            }
            
            const atual = parseFloat(parseFloat(precoAtual).toFixed(3));
            const anterior = parseFloat(parseFloat(precoAnterior).toFixed(3));
            
            if (atual > anterior) {
                return { classe: 'up', icone: 'bi bi-arrow-up' };
            } else if (atual < anterior) {
                return { classe: 'down', icone: 'bi bi-arrow-down' };
            } else {
                return { classe: 'stable', icone: '' };
            }
        }

        function obterComparacaoPrecoKGNet(produto) {
            const precoAnterior = listaUltimoPedidoFT.find(i => i.produto === produto.codigo)?.precoKGNetBlank;
            const comparacao = compararPrecoKGNet(produto.precoKGNetBlank, precoAnterior);
            const textoTooltip = comparacao.classe === 'up' ? 'Valor maior que último pedido' : comparacao.classe === 'down' ? 'Valor menor que último pedido' : '';

            return {
                classe: comparacao.classe,
                icone: comparacao.icone,
                textoTooltip: textoTooltip
            };
        }

        function obterBadgeStatusCalculo(statusCalculo) {
            switch(statusCalculo) {
                case 'Sucesso':
                    return { classe: 'badge-green', texto: 'Sucesso' };
                case 'Erro':
                    return { classe: 'badge-red', texto: 'Erro' };
                case 'Pendente':
                default:
                    return { classe: 'badge-yellow', texto: 'Pendente' };
            }
        }

        function obterAcoesProduto(produto, index) {
            const itens = [
                `<a class="dropdown-item" onclick="abrirModalResultadosCalculo('${produto.idOrcamento}')"><i class="bi bi-calculator me-1"></i> Cálculos</a>`,
                `<a class="dropdown-item" onclick="detalhesProduto('${produto.idOrcamento}', '${index}')"><i class="bi bi-eye me-1"></i> Detalhes</a>`,
            ];

            if(['CRIAR', 'EDITAR'].includes(acao)) {
                itens.push([
                    `<hr class="dropdown-divider">`,
                    `<a class="dropdown-item" onclick="editarProduto('${index}')"><i class="bi bi-pencil me-1"></i> Editar</a>`,
                    `<a class="dropdown-item" onclick="removerProduto('${index}')"><i class="bi bi-trash me-1"></i> Remover</a>`,
                ])
            }

            return `
                <td style="text-align: center">
                    <dropdown-menu label="Ações">
                        ${itens.join('')}
                    </dropdow-menu>
                </td>
            `
        }

        /* Mensagens */
        function abrirModalMensagem() {

            limparCamposMensagem();            
            const modal = new bootstrap.Modal(document.getElementById('modalMensagem'));
            modal.show();

            validacoes.modalMensagem.refresh();
        }

        function limparCamposMensagem() {
            $('#mensagemNivel').val('0');
            $('#mensagemDescricao').val('');
        }

        
        function salvarMensagem() {   
            const mensagem = {
                nomeUsuarioInclusao: '<%=usuarioLogado("nome")%>',
                numeroOrcamento: numeroOrcamento,
                nivel: $('#mensagemNivel').val(),
                descricao: $('#mensagemDescricao').val(),
                dataInclusao: formatarParaDataBR(new Date())
            };

            if (!numeroOrcamento) {
                listaMensagens.push(mensagem);
                Toast.success('Mensagem adicionada com sucesso!');
                renderizarMensagens();

                const modal = bootstrap.Modal.getInstance(document.getElementById('modalMensagem'));
                modal.hide();
                return;
            }

            exibirLoading(true);
            $.ajax({
                url: 'OrcamentoRepositorio.asp',
                type: 'POST',
                data: {
                    acao: 'salvarMensagem',
                    numeroOrcamento: numeroOrcamento,
                    nivel: mensagem.nivel,
                    descricao: mensagem.descricao
                },
                dataType: 'json',
                success: function(response) {
                    if (!response.sucesso) {
                        Toast.error(response.mensagem || 'Erro ao salvar mensagem');
                        return;
                    }
                    
                    listaMensagens.push(response.dados || mensagem);
                    Toast.success(response.mensagem || 'Mensagem adicionada com sucesso!');
                    renderizarMensagens();

                    const modal = bootstrap.Modal.getInstance(document.getElementById('modalMensagem'));
                    modal.hide();
                },
                error: function(xhr, status, error) {
                    console.error('Erro ao salvar mensagem:', error);
                    Toast.error('Erro ao salvar mensagem');
                },
                complete: function() {
                    exibirLoading(false);
                }
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

        function normalizarMensagemDetalhe(valor) {
            const texto = String(valor || '').trim();
            return texto || '-';
        }

        function criarCelulaDetalheMensagem(valor) {
            const texto = normalizarMensagemDetalhe(valor);
            const textoSeguro = escapeHtml(texto);

            if (texto === '-') {
                return '<td class="mensagem-detalhe-cell text-muted">-</td>';
            }

            return `
                <td class="mensagem-detalhe-cell">
                    <div class="mensagem-truncate-text mensagem-detalhe-link"
                         role="button"
                         tabindex="0"
                         onclick="abrirModalDetalheMensagem(this)"
                         onkeydown="abrirModalDetalheMensagemTeclado(event, this)"
                         data-detalhe="${textoSeguro}">
                        ${textoSeguro}
                    </div>
                </td>
            `;
        }

        function abrirModalDetalheMensagem(elemento) {
            const detalhe = $(elemento).attr('data-detalhe') || '-';
            $('#modalDetalheMensagemConteudo').text(detalhe);
            const modal = new bootstrap.Modal(document.getElementById('modalDetalheMensagem'));
            modal.show();
        }

        function abrirModalDetalheMensagemTeclado(event, elemento) {
            if (event.key === 'Enter' || event.key === ' ') {
                event.preventDefault();
                abrirModalDetalheMensagem(elemento);
            }
        }

        
        function renderizarMensagens() {
            const $container = $('#mensagemListaContainer');
            const $contador = $('#contadorMensagens, #menuContadorMensagens');
            
            $contador.text(listaMensagens.length);

            if (listaMensagens.length === 0) {
                $container.html(`
                    <div class="text-center py-5">
                        <p class="text-muted mt-3 mb-0">Nenhuma mensagem adicionada.</p>
                    </div>
                `);
                return;
            }

            let html = `
                <div class="table-responsive">
                    <table id="gradeMensagens" class="table table-hover table-sm mb-0">
                        <thead class="table-light">
                            <tr>
                                <th>Usuário</th>
                                <th>Mensagem</th>
                                <th>Data Inclusao</th>
                            </tr>
                        </thead>
                        <tbody>
            `;
            
            listaMensagens.forEach((item, index) => {
                html += `
                    <tr>
                        <td>${item.nomeUsuarioInclusao || '-'}</td>
                        ${criarCelulaDetalheMensagem(item.descricao)}
                        <td>${item.dataInclusao}</td>
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

        /* Cálculo */
        function abrirModalCalculo() {
            if(produtosSelecionados.size === 0) {
                Toast.warning('Nenhum produto selecionado');
                return;
            }

            limparCamposCalculo();

            const modal = new bootstrap.Modal(document.getElementById('modalCalculo'));
            modal.show();
        }

        function toggleSelecao(idOrcamento) {
            idOrcamento = String(idOrcamento);
            var produto = listaProdutos.find(function(p) { return String(p.idOrcamento) === idOrcamento; });
            
            if (!produto) {
                produto = produtosSelecionados.get(idOrcamento);
            }
            
            if (!produto) return;
            
            if (produtosSelecionados.has(idOrcamento)) {
                produtosSelecionados.delete(idOrcamento);
            } else {
                produtosSelecionados.set(idOrcamento, produto);
            }

            atualizarCheckboxSelectAll();
        }
            
        function toggleSelectAll() {
            var checkbox = $('#checkboxSelectAll');
            var selecionar = checkbox.is(':checked');
            
            if (selecionar) {
                listaProdutos.forEach(function(produto) {
                    produtosSelecionados.set(String(produto.idOrcamento), produto);
                });
            } else {
                produtosSelecionados.clear();
            }

            renderizarProdutos();
        }
            
        function atualizarCheckboxSelectAll() {
            var checkbox = $('#checkboxSelectAll');
            
            if (listaProdutos.length === 0) {
                checkbox.prop('checked', false);
                checkbox.prop('indeterminate', false);
                return;
            }
            
            var todosSelecionados = listaProdutos.every(function(produto) {
                return produtosSelecionados.has(String(produto.idOrcamento));
            });
            
            var algunsSelecionados = listaProdutos.some(function(produto) {
                return produtosSelecionados.has(String(produto.idOrcamento));
            });

            if (todosSelecionados) {
                checkbox.prop('checked', true);
                checkbox.prop('indeterminate', false);
            } else if (algunsSelecionados) {
                checkbox.prop('checked', false);
                checkbox.prop('indeterminate', true);
            } else {
                checkbox.prop('checked', false);
                checkbox.prop('indeterminate', false);
            }
        }
            
        function realizarCalculos() {
            if (produtosSelecionados.size === 0) {
                Toast.warning('Nenhum orcamento selecionado');
                return;
            }

            const codPlanta = $('#codPlanta').val();
            const tipoCalculo = $('#calculoTipo').val();
            const valorBase = $('#calculoValorBase').val() || 0;
            const idsOrcamento = Array.from(produtosSelecionados.keys());

            const payload = {
                acao: 'calcularTrim',
                codPlanta: codPlanta,
                tipoCalculo: tipoCalculo,
                valorBase: valorBase,
                itemsIds: (idsOrcamento.map(function(idOrcamento) { return "'" + idOrcamento.trim() + "'"; }).join(',')),
            }

            exibirLoading(true);
            const modal = bootstrap.Modal.getInstance(document.getElementById('modalCalculo'));
            modal.hide();

            $.ajax({
                url: 'OrcamentoCalculoTrimboxRepositorio.asp',
                type: 'POST',
                dataType: 'json',
                data: payload,
                success: function(response) {
                    if (!response.sucesso) {
                        Toast.error(response.mensagem || 'Erro ao realizar cálculos');
                        return;
                    }
                    
                    resultadosCalculo.sucessoIds = [...resultadosCalculo.sucessoIds, ...response.sucessoIds]
                    resultadosCalculo.erroIds = [ ...resultadosCalculo.erroIds, ...response.erroIds]
                    produtosSelecionados.clear();

                    exibirLoading(false);
                    carregarItens();
                    abrirModalResultadosCalculo();
                },
                error: function(xhr, status, error) {
                    console.error('Erro ao realizar cálculos:', error);
                    Toast.error('Erro ao realizar cálculos');
                    exibirLoading(false);
                }
            });
        }

        function abrirModalResultadosCalculo(idOrcamento = "") {
            const modal = new bootstrap.Modal(document.getElementById('modalResultadosCalculo'));
            modal.show();
            
            $("#resultadosCalculo").html('<div class="text-center p-4"> <div class="spinner-border text-primary mb-3" role="status"></div></div>');

            $.ajax({
                url: 'OrcamentoCalculoTrimboxResultados.asp?v=' + Math.random(),
                type: 'POST',
                data: { 
                    idOrcamentoSknet: idOrcamento, 
                    sucessoIds: resultadosCalculo.sucessoIds.join(','),
                    erroIds: resultadosCalculo.erroIds.join(',')
                },
                success: function(response) {
                    $('#resultadosCalculo').html(response);
                    resultadosCalculo.sucessoIds = [];
                    resultadosCalculo.erroIds = [];
                },
                error: function(xhr, status, error) {
                    $('#resultadosCalculo').html('<div class="alert alert-danger">Erro ao carregar resultados: ' + error + '</div>');
                    resultadosCalculo.sucessoIds = [];
                    resultadosCalculo.erroIds = [];
                }
            });
        }

        function abrirModalUltimoPedidoFT() {
            const modal = new bootstrap.Modal(document.getElementById('modalUltimoPedidoFT'));
            modal.show();

            const $tbody = $('#gradeUltimoPedidoFT tbody');
            $tbody.empty();
     
            if (listaUltimoPedidoFT.length === 0) {
                $tbody.html(`
                    <tr>
                        <td colspan="28" class="text-center text-muted py-4">
                            Nenhum registro encontrado
                        </td>
                    </tr>
                `);
                return;
            }

            listaUltimoPedidoFT.forEach(function(item) {    
                $tbody.append(`
                    <tr>
                        <td>${item.produto || '-'}</td>
                        <td>${item.referencia || '-'}</td>
                        <td>${item.pedido || '-'}</td>
                        <td>${item.precoUnitario || '-'}</td>
                        <td>${item.condicaoPagamento || '-'}</td>
                        <td>${item.precoKGNetBlank || '-'}</td>
                        <td>${item.dataEmissao || '-'}</td>
                    </tr>
                `);
            });
        }


        /**
         * Auxiliar para limpar campos de produto
         * Se não for produto novo, desabilita os campos que não podem ser editados
         */
        function limparCamposCalculo() {            
            $('#calculoTipo').val('VU').trigger('change.select2');
            $('#calculoValorBase').val('0');
        }

        /** 
         * Auxiliar para manter o resultado do calculo de área e peso sem arredondar (regra de negócio) 
        */
        function trucarValor(num, decimals) {
            const regex = new RegExp(`^-?\\d+(?:\\.\\d{0,${decimals}})?`);
            return parseFloat(num.toString().match(regex)[0]);
        }

        /* Auxiliares para validações */
        const naoVazio = (msg) => ({
            validator: (value) => value !== null && value !== undefined && String(value).trim() !== '',
            errorMessage: msg
        });

        const maiorQueZero = (msg) => ({
            validator: (value) => value !== null && value !== undefined && parseFloat(value) > 0,
            errorMessage: msg
        })

        const obrigatorioQuando = (condicao, msg) => ({
            validator: (value) => !condicao() || (value !== null && value !== undefined && String(value).trim() !== ''),
            errorMessage: msg
        })

        const maiorQueZeroQuando = (condicao, msg) => ({
            validator: (value) => {
                if (!condicao()) return true;
                if (value === null || value === undefined || String(value).trim() === '') return true;
                return parseFloat(value) > 0;
            },
            errorMessage: msg
        })


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
            const tipoCaixa = $('#produtoTipoCaixa').val();
            const produtoNaoChapa = tipoCaixa != eTipoCaixa.CHAPAS;
            const produtoTipoCaixaCorteVincoOuChapas = tipoCaixa == eTipoCaixa.CORTE_VINCO || tipoCaixa == eTipoCaixa.CHAPAS;
            const produtoTipoCaixaCorteVinco = tipoCaixa == eTipoCaixa.CORTE_VINCO;
            const produtoTemIPI = $('#produtoTemIPI').val() == '1';
            const alturaObrigatoria = tipoCaixa != eTipoCaixa.CHAPAS && tipoCaixa != eTipoCaixa.CORTE_VINCO;

            const camposSempreObrigatorios = [
                'produtoTipoItem', 'produtoTipoCaixa', 'produtoCod', 'produtoEstrutura',
                'produtoReferencia', 'produtoQuantidade', 'produtoTipoPreco', 'produtoValorUnitario',
                'produtoGramatura', 'produtoTemIPI'
            ];

            $('#formularioProduto .required-field').hide();

            camposSempreObrigatorios.forEach(campo => definirMarcacaoObrigatoria(campo, true));

            definirMarcacaoObrigatoria('produtoComprimentoInterno', produtoNaoChapa);
            definirMarcacaoObrigatoria('produtoLarguraInterna', produtoNaoChapa);
            definirMarcacaoObrigatoria('produtoAlturaInterna', alturaObrigatoria);
            definirMarcacaoObrigatoria('produtoComprimentoChapa', produtoTipoCaixaCorteVincoOuChapas);
            definirMarcacaoObrigatoria('produtoLarguraChapa', produtoTipoCaixaCorteVincoOuChapas);
            definirMarcacaoObrigatoria('produtoAreaRefile', produtoTipoCaixaCorteVincoOuChapas);
            definirMarcacaoObrigatoria('produtoPesoRefile', produtoTipoCaixaCorteVincoOuChapas);
            definirMarcacaoObrigatoria('produtoArranjo', produtoTipoCaixaCorteVinco);
            definirMarcacaoObrigatoria('produtoAliquotaIPI', produtoTemIPI);
        }

        function inicializarValidacoes() {
            // Aba 1
            validacoes.aba1 = new JustValidate('#formularioOrcamento', {
                lockForm: false, 
            });

            validacoes.aba1
                .addField('#codPlanta', [naoVazio('Planta é obrigatória')])
                .addField('#codCliente', [naoVazio('Cliente é obrigatório')])
                .addField('#codLoja', [naoVazio('Loja é obrigatória')])
                .addField('#nomeCliente', [naoVazio('Nome do cliente é obrigatório')])
                .addField('#codVendedor', [naoVazio('Cód. do vendedor é obrigatório')])
                .addField('#nomeVendedor', [naoVazio('Nome do vendedor é obrigatório')])
                .addField('#CEP', [naoVazio('CEP é obrigatório')])
                .addField('#endereco', [naoVazio('Endereço é obrigatório')])
                .addField('#bairro', [naoVazio('Bairro é obrigatório')])
                .addField('#UF', [naoVazio('UF é obrigatória')])
                .addField('#municipioId', [naoVazio('Município é obrigatório')])

            validacoes.aba2 = new JustValidate('#formularioOrcamento', {
                lockForm: false, 
            });

            validacoes.aba2
                .addField('#dataValidadeDe', [naoVazio('Data de validade de é obrigatória')])
                .addField('#dataValidadeAte', [naoVazio('Data de validade até é obrigatória')])
                .addField('#tipoFrete', [naoVazio('Tipo de frete é obrigatório')])
                .addField('#tipoCaminhao', [naoVazio('Tipo de caminhão é obrigatório')])
                .addField('#condicaoPagamentoId', [naoVazio('Condição de pagamento é obrigatória')])
                .addField('#temICMS', [naoVazio('ICMS é obrigatório')])
                .addField('#aliquotaICMS', [naoVazio('Aliquota ICMS é obrigatória')])
                .addField('#temPIS', [naoVazio('PIS é obrigatório')])
                .addField('#temCOFINS', [naoVazio('COFINS é obrigatório')]);

            // Validação é feita de forma manual
            validacoes.aba3 = new JustValidate('#formularioOrcamento', {
                lockForm: false, 
            });

            // Sem campos obrigatórios
            validacoes.aba4 = new JustValidate('#formularioOrcamento', {
                lockForm: false, 
            });

            validacoes.modalProduto = new JustValidate('#formularioProduto', {
                lockForm: false, 
            });
            
            const produtoTipoCaixaCorteVincoOuChapas = () => {
                const tipoCaixa = $('#produtoTipoCaixa').val();
                return tipoCaixa == eTipoCaixa.CORTE_VINCO || tipoCaixa == eTipoCaixa.CHAPAS;
            }

            const produtoNaoChapa = () => $('#produtoTipoCaixa').val() != eTipoCaixa.CHAPAS;
            const produtoTipoCaixaCorteVinco = () => $('#produtoTipoCaixa').val() == eTipoCaixa.CORTE_VINCO;
            const produtoTipoCaixaNaoChapasCorteVinco = () => $('#produtoTipoCaixa').val() != eTipoCaixa.CHAPAS && $('#produtoTipoCaixa').val() != eTipoCaixa.CORTE_VINCO;
            const produtoTemIPI = () => $('#produtoTemIPI').val() == '1';

            validacoes.modalProduto
                .addField('#produtoTipoItem', [naoVazio('Tipo do item é obrigatório')])
                .addField('#produtoTipoCaixa', [naoVazio('Tipo de caixa é obrigatória')])
                .addField('#produtoCod', [naoVazio('Código do produto é obrigatório')])
                .addField('#produtoEstrutura', [naoVazio('Estrutura é obrigatória')])
                .addField('#produtoReferencia', [naoVazio('Referência é obrigatória')])
                .addField('#produtoTipoPreco', [naoVazio('Tipo de preço é obrigatório')])
                .addField('#produtoQuantidade', [maiorQueZero('Quantidade deve ser maior que zero')])
                .addField('#produtoValorUnitario', [maiorQueZero('Valor unitário deve ser maior que zero')])
                .addField('#produtoGramatura', [maiorQueZero('Gramatura deve ser maior que zero')])
                .addField('#produtoComprimentoInterno', [
                    obrigatorioQuando(produtoNaoChapa, 'Comprimento interno é obrigatório'),
                    maiorQueZeroQuando(produtoNaoChapa, 'Comprimento interno deve ser maior que zero')
                ])
                .addField('#produtoLarguraInterna', [
                    obrigatorioQuando(produtoNaoChapa, 'Largura interno é obrigatória'),
                    maiorQueZeroQuando(produtoNaoChapa, 'Largura interna deve ser maior que zero')
                ])
                .addField('#produtoAlturaInterna', [
                    obrigatorioQuando(produtoTipoCaixaNaoChapasCorteVinco, 'Altura interna é obrigatória'),
                    maiorQueZeroQuando(produtoTipoCaixaNaoChapasCorteVinco, 'Altura interna deve ser maior que zero')
                ])
                .addField('#produtoComprimentoChapa', [
                    obrigatorioQuando(produtoTipoCaixaCorteVincoOuChapas, 'Comprimento da chapa é obrigatório'),
                    maiorQueZeroQuando(produtoTipoCaixaCorteVincoOuChapas, 'Comprimento da chapa deve ser maior que zero')
                ])
                .addField('#produtoLarguraChapa', [
                    obrigatorioQuando(produtoTipoCaixaCorteVincoOuChapas, 'Largura da chapa é obrigatória'),
                    maiorQueZeroQuando(produtoTipoCaixaCorteVincoOuChapas, 'Largura da chapa deve ser maior que zero')
                ])
                .addField('#produtoArranjo', [
                    obrigatorioQuando(produtoTipoCaixaCorteVinco, 'Arranjo é obrigatório'),
                    maiorQueZeroQuando(produtoTipoCaixaCorteVinco, 'Arranjo deve ser maior que zero')
                ])
                .addField('#produtoAreaRefile', [
                    obrigatorioQuando(produtoTipoCaixaCorteVincoOuChapas, 'Área refile é obrigatória'),
                    maiorQueZeroQuando(produtoTipoCaixaCorteVincoOuChapas, 'Área refile deve ser maior que zero')
                ])
                .addField('#produtoPesoRefile', [
                    obrigatorioQuando(produtoTipoCaixaCorteVincoOuChapas, 'Peso refile é obrigatório'),
                    maiorQueZeroQuando(produtoTipoCaixaCorteVincoOuChapas, 'Peso refile deve ser maior que zero')
                ])
                .addField('#produtoTemIPI', [naoVazio('IPI é obrigatório')])
                .addField('#produtoAliquotaIPI', [
                    obrigatorioQuando(produtoTemIPI, 'Aliquota IPI é obrigatória'),
                    maiorQueZeroQuando(produtoTemIPI, 'Aliquota IPI deve ser maior que zero')
                ])
                
            validacoes.modalMensagem = new JustValidate('#formularioMensagem', {
                lockForm: false, 
            });

            validacoes.modalMensagem
                .addField('#mensagemNivel', [naoVazio('Para quem deseja enviar é obrigatório')])
                .addField('#mensagemDescricao', [naoVazio('Mensagem é obrigatória')]);

            atualizarMarcacoesObrigatorias();
        }

        function navegarParaGrade() {
            window.location.href = `OrcamentoGrade.asp`;
        }

        function navegarParaAba(numeroAba) {
            if (numeroAba < 1 || numeroAba > totalAbas) return;

            abaAtual = numeroAba;

            $('.tab-button').removeClass('active');
            $(`.tab-button[data-tab="${numeroAba}"]`).addClass('active');

            $('.tab-content').removeClass('active');
            $(`.tab-content[data-tab-content="${numeroAba}"]`).addClass('active');
        }

        function irProximo() {
            navegarParaAba(abaAtual + 1);
        }

        function irAnterior() {
            navegarParaAba(abaAtual - 1);
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

        function carregarDadosOrcamento() {
            exibirLoading(true);

            const params = { 
                acao: 'obterPorNumeroOrcamento', 
                numeroOrcamento: numeroOrcamento, 
            }

            $.ajax({
                url: 'OrcamentoRepositorio.asp',
                type: 'GET',
                data: params,
                dataType: 'json',
                success: function(response) {
                    if (!response.sucesso) {
                        Toast.error(response.mensagem);
                        return;
                    }

                    statusAtual = response.dados.status;
                    preencherFormulario(response.dados);
                },
                error: function(xhr, status, error) {
                    console.error('Erro ao carregar dados para edição:', error);
                    Toast.error('Erro ao carregar dados para edição');
                }, 
            });
        }

        function carregarUltimoPedidoFT(codPlanta, codCliente, codLoja, listaProdutos) {
            const params = { 
                acao: 'obterUltimoPedidoFT', 
                codPlanta: codPlanta,
                codCliente: codCliente,
                codLoja: codLoja,
                codProdutos: (listaProdutos.map(function(produto) { return "'" + produto.codigo + "'"; }).join(',')),
            }

            return $.ajax({
                url: 'OrcamentoRepositorio.asp',
                type: 'POST',
                data: params,
                dataType: 'json',
                success: function(response) {
                    if (!response.sucesso) {
                        Toast.error(response.mensagem);
                        return;
                    }

                    listaUltimoPedidoFT = response.dados || [];
                },
                error: function(xhr, status, error) {
                    console.error('Erro ao carregar dados para edição:', error);
                    Toast.error('Erro ao carregar dados para edição');
                }
            });
        }

        function carregarItens() {
            const params = { 
                acao: 'listarItens', 
                numeroOrcamento: numeroOrcamento, 
            }

            $.ajax({
                url: 'OrcamentoRepositorio.asp',
                type: 'GET',
                data: params,
                dataType: 'json',
                success: function(response) {
                    if (!response.sucesso) {
                        Toast.error(response.mensagem);
                        return;
                    }

                    listaProdutos = response.dados || [];
                    renderizarProdutos();
                },
                error: function(xhr, status, error) {
                    console.error('Erro ao carregar lista de itens:', error);
                    Toast.error('Erro ao carregar lista de itens');
                }
            });
        }

        function preencherFormulario(dados) {
            var p1 = carregarMunicipios(dados.codPlanta, dados.codMunicipio, dados.UF);
            var p2 = carregarCondicoesPagamento(dados.codPlanta);
            var p3 = carregarTiposCaminhao(dados.codPlanta);
            var p4 = carregarEstruturas(dados.codPlanta);
            var p5 = carregarParametrosProduto(dados.codPlanta);
            var p6 = carregarUltimoPedidoFT(dados.codPlanta, dados.codCliente, dados.codLoja, dados.itens);

            $.when(p1, p2, p3, p4, p5, p6).then(function() {
                preencherDadosCliente(dados, false, false);
                $('#dataValidadeDe').val(dados.dataValidadeDe)
                $('#dataValidadeAte').val(dados.dataValidadeAte)
                $('#aliquotaICMS').val(dados.aliquotaICMS)

                listaProdutos = dados.itens || [];
                listaMensagens = dados.mensagens || [];
                renderizarProdutos(); 
                renderizarMensagens();
  
                // Campos select precisam de um pequeno delay
                setTimeout(() => {
                    $('#codPlanta').val(dados.codPlanta).trigger('change.select2')
                    $('#UF').val(dados.UF).trigger('change.select2')
                    $('#municipioId').val(dados.municipioId).trigger('change.select2')
                    $('#temPIS').val(dados.temPIS).trigger('change.select2')
                    $('#temCOFINS').val(dados.temCOFINS).trigger('change.select2')
                    $('#tipoFrete').val(dados.tipoFrete).trigger('change.select2')
                    $('#tipoCaminhao').val(dados.tipoCaminhao).trigger('change.select2')
                    $('#condicaoPagamentoId').val(dados.condicaoPagamentoId).trigger('change.select2')
                    $('#temICMS').val(dados.temICMS).trigger('change.select2')

                    exibirLoading(false);
                }, 500);
            });
        }

        function anexarEventos() {
            /* Formulário principal */
            $('#codPlanta').on('change', function() {
                campoPlantaAlterado();
            });

            $('#UF').on('change', function() {
                campoUFAlterado();
            });

            $('#temICMS').on('change', function() {
                campoTemICMSAlterado();
            });

            $('#btnSalvar').on('click', async function() {
                validarFormulario();
            });

            /* Formulário (modal) de produto */
            $('#produtoTipoItem').on('change', function() {
                campoProdutoTipoAlterado();
            });
            
            $('#produtoTipoCaixa').on('change', function() {
                campoProdutoTipoCaixaAlterado();
            });

            $('#produtoTemIPI').on('change', function() {
                campoProdutoTemIPIAlterado();
            });

            $('#produtoTipoPreco').on('change', function() {
                atualizarLabelValorUnitario();
            });

            $('#produtoEstrutura').on('change', function() {
                campoProdutoEstruturaAlterado();
            });

            $('#produtoComprimentoChapa, #produtoLarguraChapa, #produtoArranjo, #produtoPecasPorConjunto').on('input', function() {
                calcularAreaRefile();
            });

            $('#produtoGramatura, #produtoAreaRefile').on('input', function() {
                calcularPesoRefile();
            });

            $('#flagProdutoNovo').on('change', function() {
                limparCamposProduto();
            });

            $('#flagClienteNovo').on('change', function() {
                limparCamposCliente();
            });

            $('#produtoQuantidade').on('beforeinput', function(e) {
                const data = e.originalEvent.data;
                if (data && /[^0-9]/.test(data)) {
                    e.preventDefault();
                }
            });
    
            $('#codCliente').on('input', function() {
                if (!$(this).val().trim()) {
                    limparCamposCliente();
                }
            });

            $('#produtoCod').on('input', function() {
                if (!$(this).val().trim()) {
                    limparCamposProduto();
                }
            });
    
            /* Navegação */
            $('#btnVoltar').on('click', function() { 
                navegarParaAba(abaAtual - 1); 
            });

            $('#btnProximo').on('click', function() {
                // Navegação livre entre abas; validação apenas no envio do formulário.
                navegarParaAba(abaAtual + 1);
            });

            $('.tab-button').on('click', function() {
                const tabNumber = parseInt($(this).data('tab'));

                if (tabNumber === abaAtual) {
                    return;
                }

                navegarParaAba(tabNumber);
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

            /* Modal de buscar vendedor */
            let buscaVendedorTimeout = null;
            $('#vendedorBusca').on('input', function() {
                clearTimeout(buscaVendedorTimeout);

                buscaVendedorTimeout = setTimeout(function() {
                    pesquisarVendedor();
                }, 500);
            });

            $('#vendedorBusca').on('keypress', function(e) {
                if (e.which === 13) {
                    e.preventDefault();
                    clearTimeout(buscaVendedorTimeout);
                    pesquisarVendedor();
                }
            });

            /* Modal de buscar produto */
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

            /* Botões de ações */
            $('#btnEnviarParaAprovacao').on('click', async function () {
                enviarParaAprovacao();
            })

            $('#btnDevolver').on('click', async function () {
                devolver();
            })

            $('#btnSalvarProduto').on('click', function () {
                const validador = validacoes.modalProduto;

                validador.revalidate().then(isValid => {
                    if (!isValid) {
                        const errorsParaExibir = {};
                        Object.entries(validador.fields).forEach(([index, campo]) => {
                            if(campo.isValid) return;

                            const primeiroErro = campo.rules.find(rule => !rule.isValid);
                            if (primeiroErro) {
                                errorsParaExibir[`#${campo.elem.id}`] = primeiroErro.errorMessage || 'Campo inválido';
                            }   
                        });

                        validador.showErrors(errorsParaExibir);
                        return;
                    }

                    salvarProduto();
                });
            });

            $('#btnSalvarMensagem').on('click', function () {
                const validador = validacoes.modalMensagem;

                validador.revalidate().then(isValid => {
                    if (!isValid) {
                        const errorsParaExibir = {};
                        Object.entries(validador.fields).forEach(([index, campo]) => {
                            if(campo.isValid) return;

                            const primeiroErro = campo.rules.find(rule => !rule.isValid);
                            if (primeiroErro) {
                                errorsParaExibir[`#${campo.elem.id}`] = primeiroErro.errorMessage || 'Campo inválido';
                            }   
                        });

                        validador.showErrors(errorsParaExibir);
                        return;
                    }

                    salvarMensagem();
                });
            });

            $('#btnConfirmarCalculo').on('click', function () {
                realizarCalculos();
            });

            $('#btnUltimoPedidoFT').on('click', function () {
                abrirModalUltimoPedidoFT();
            });
        }

        async function validarFormulario() {
            const abas = [1, 2, 3, 4];
            const resultados = await Promise.all([
                validacoes.aba1.revalidate(),
                validacoes.aba2.revalidate(),
                Promise.resolve(listaProdutos.length > 0),
                validacoes.aba4.revalidate(),
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

            if (!resultados[2]) {
                Toast.error('Adicione ao menos um item.');
            }

            if (!resultados.every(Boolean)) {
                const primeiraInvalida = abas[resultados.findIndex(v => !v)];
                navegarParaAba(primeiraInvalida);
                Toast.error('Existem campos obrigatórios pendentes. Verifique as seções destacadas.');
                return { sucesso: false };
            }

            return enviarFormulario();
        }

        
        function enviarParaAprovacao() {
            exibirLoading(true)

            const dadosFormulario = {
                acao:'enviarParaAprovacao',
                numeroOrcamento: numeroOrcamento,
            };

            return $.ajax({
                url: 'OrcamentoRepositorio.asp',
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
                    }, 1500)
                },
                error: function(xhr, status, error) {
                    Toast.error('Erro ao salvar formulário');
                },
                complete: function() {
                    exibirLoading(false)
                }
            });
            
        }

        function devolver() {
            exibirLoading(true)

            const dadosFormulario = {
                acao:'devolver',
                numeroOrcamento: numeroOrcamento,
            };

            return $.ajax({
                url: 'OrcamentoRepositorio.asp',
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
                    }, 1500)
                },
                error: function(xhr, status, error) {
                    Toast.error('Erro ao salvar formulário');
                },
                complete: function() {
                    exibirLoading(false)
                }
            });
            
        }

        function enviarFormulario() {
            exibirLoading(true)

            const dadosFormulario = {
                flagClienteNovo: $('#flagClienteNovo').is(':checked') ? '1' : '0',
                status: statusAtual,
                codPlanta: $('#codPlanta').val(),
                codCliente: $('#codCliente').val(),
                codLoja: $('#codLoja').val(),
                nomeCliente: $('#nomeCliente').val().replace(/'/g, ""),
                CNPJ: $('#CNPJ').val().replace(/[^a-zA-Z0-9\s]/g, ''),
                codVendedor: $('#codVendedor').val(),
                nomeVendedor: $('#nomeVendedor').val().replace(/'/g, ""),
                telefone: $('#telefone').val().replace(/[^0-9]/g, ''),
                contato: $('#contato').val(),
                email: $('#email').val(),
                CEP: $('#CEP').val().replace(/[^0-9]/g, ''),
                endereco: $('#endereco').val().replace(/'/g, ""),
                numero: $('#numero').val().replace(/[^0-9]/g, ''),
                complemento: $('#complemento').val().replace(/'/g, ""),
                bairro: $('#bairro').val().replace(/'/g, ""),
                UF: $('#UF').val(),
                municipioId: $('#municipioId').val(),
                nomeMunicipio: $('#municipioId').find('option:selected').text().replace(/'/g, ""),
                dataValidadeDe: $('#dataValidadeDe').val(),
                dataValidadeAte: $('#dataValidadeAte').val(),
                tipoFrete: $('#tipoFrete').val(),
                tipoCaminhao: $('#tipoCaminhao').val(),
                condicaoPagamentoId: $('#condicaoPagamentoId').val(),
                temICMS: $('#temICMS').val(),
                aliquotaICMS: $('#aliquotaICMS').val(),
                temPIS: $('#temPIS').val(),
                temCOFINS: $('#temCOFINS').val(),
                itens: JSON.stringify(listaProdutos),
            };

            if (numeroOrcamento) {
                dadosFormulario.numeroOrcamento = numeroOrcamento;
                dadosFormulario.acao = 'processar';
            } else {
                dadosFormulario.acao = 'criar';
                dadosFormulario.mensagens = JSON.stringify(listaMensagens);
            }

            $.ajax({
                url: 'OrcamentoRepositorio.asp',
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
                        //navegarParaGrade();
                        window.location.href = `OrcamentoFormulario.asp?acao=EDITAR&numeroOrcamento=${response.numeroOrcamento}`;
                    }, 1500)
                },
                error: function(xhr, status, error) {
                    Toast.error('Erro ao salvar formulário');
                },
                complete: function() {
                    exibirLoading(false)
                }
            });
        }

        const configuracoesPorAcao = Object.freeze({
            CRIAR: {
                titulo: () => 'Cadastro de Orcamento',
                camposDesabilitados: '',
                camposHabilitados: '#formularioOrcamento input, #formularioOrcamento textarea, #formularioOrcamento select',
                botoesVisiveis: '#btnSalvar, #btnAdicionarProduto, #btnAdicionarMensagem, #avisoCalculoProdutos'
            },
            EDITAR: {
                titulo: () => `Editando Orcamento #${numeroOrcamento}`,
                camposDesabilitados: `#flagClienteNovo, #codPlanta, #btnBuscarCliente, #codCliente, #codLoja,
                    #nomeCliente, #CNPJ, #codVendedor, #nomeVendedor, #email, #CEP, #endereco,
                    #bairro, #UF, #municipioId, #numero, #telefone, #contato, #complemento`,
                camposHabilitados: `#dataValidadeDe, #dataValidadeAte, #aliquotaICMS, #temPIS,
                    #temCOFINS, #tipoFrete, #tipoCaminhao, #condicaoPagamentoId, #temICMS`,
                botoesVisiveis: '#btnSalvar, #btnAdicionarProduto, #btnAdicionarMensagem, #avisoCalculoProdutos'
            },
            VISUALIZAR: {
                titulo: () => `Visualizando Orcamento #${numeroOrcamento}`,
                camposDesabilitados: '#btnBuscarCliente, #formularioOrcamento input, #formularioOrcamento textarea, #formularioOrcamento select',
                camposHabilitados: '',
                botoesVisiveis: ''
            },
            ANALISAR: {
                titulo: () => `Analisando Orcamento #${numeroOrcamento}`,
                camposDesabilitados: '#btnBuscarCliente, #formularioOrcamento input, #formularioOrcamento textarea, #formularioOrcamento select',
                camposHabilitados: '',
                botoesVisiveis: '#btnDevolver, #btnAdicionarMensagem, #btnUltimoPedidoFT, #btnEnviarParaAprovacao, #btnCalcular'
            }
        });

        function normalizarAcao(valor) {
            return String(valor || 'CRIAR').trim().toUpperCase();
        }

        function habilitarCampos(seletor, habilitar) {
            if (!seletor) return;
            $(seletor).prop('disabled', !habilitar);
            $(seletor).filter('select').trigger('change.select2');
        }


        function aplicarModoPorAcao() {
            acao = normalizarAcao(acao);
            if (!configuracoesPorAcao[acao]) acao = 'CRIAR';
            const configuracao = configuracoesPorAcao[acao];

            $('#titulo').text(configuracao.titulo());
            habilitarCampos(configuracao.camposDesabilitados, false);
            habilitarCampos(configuracao.camposHabilitados, true);

            if (acao === 'CRIAR') {
                aplicarEstadoCamposCliente();
            }

            if (configuracao.botoesVisiveis) {
                $(configuracao.botoesVisiveis).removeClass('d-none');
            }

            renderizarProdutos();
        }


        $(document).ready(async function() {
            const urlParams = new URLSearchParams(window.location.search);
            const paramNumeroOrcamento = urlParams.get('numeroOrcamento');
            const paramAcao = urlParams.get('acao');    
            numeroOrcamento = paramNumeroOrcamento;
            acao = paramAcao;

            if (numeroOrcamento) {
                navegarParaAba(3)
                carregarDadosOrcamento();
            }

            aplicarModoPorAcao();
            inicializarValidacoes();
            anexarEventos();
            inicializarSelect2();
            inicializarMascaras();
        });
    </script>
</head>
<body>
    <div class="container my-4">
        <div class="d-flex flex-column flex-md-row justify-content-between align-items-md-center gap-3 my-4">
            <div class="d-flex align-items-center">
                <button onclick="navegarParaGrade()" class="btn btn-sm btn-light me-2 fs-5"> 
                    <i class="bi bi-arrow-left"></i> 
                </button>
                <div>
                    <h3 id="titulo" class="mb-0"></h3>
                </div>
            </div>

            <div class="flex-wrap justify-content-end gap-2">
                <button type="button" class="d-none btn btn-sm btn-outline-secondary" id="btnDevolver">
                    <i class="bi bi-arrow-counterclockwise me-1"></i>
                    <span>Devolver</span>
                </button>
                <button type="button" class="d-none btn btn-sm btn-outline-primary" id="btnEnviarParaAprovacao">
                    <i class="bi bi-shield-check me-1"></i>
                    <span>Enviar Para Aprovação</span>
                </button>
                <button class="d-none btn btn-sm btn-success" id="btnSalvar">
                    <i class="bi bi-check-lg me-1"></i>
                    Salvar alterações
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
                            <i class="tab-icon bi bi-currency-dollar"></i>
                            <span class="tab-title">Impostos e Condições</span>
                        </button>
                        <div class="tab-group-title">Itens</div>
                        <button type="button" class="tab-button" data-tab="3">
                            <i class="tab-icon bi bi-box"></i>
                            <span class="tab-title">Produtos</span>
                            <span class="tab-meta" id="menuContadorProdutos">0</span>
                        </button>
                        <div class="tab-group-title">Histórico</div>
                        <button type="button" class="tab-button" data-tab="4">
                            <i class="tab-icon bi bi-envelope"></i>
                            <span class="tab-title">Mensagens</span>
                            <span class="tab-meta" id="menuContadorMensagens">0</span>
                        </button>
                    </div>
                </div>

                <div class="tab-content-wrapper">
                    <form id="formularioOrcamento" name="formularioOrcamento" novalidate>
                        <input type="hidden" name="hidCodMunicipio" id="hidCodMunicipio" value="" />
                       
                        <!-- Campos para calculo de quantidade mínima ao adicionar produtos -->
                        <input type="hidden" name="hidAreaMinimaMaleta" id="hidAreaMinimaMaleta" value="0" />
                        <input type="hidden" name="hidAreaMinimaCorteVinco" id="hidAreaMinimaCorteVinco" value="0" />
                        <input type="hidden" name="hidAreaMinimaChapa" id="hidAreaMinimaChapa" value="0" />

                        <!-- Aba Dados Cadastrais -->
                        <div class="tab-content active" data-tab-content="1">
                            <div class="section-header">
                                <div class="section-title">Cadastro do orçamento</div>
                                <p class="section-subtitle">
                                    Dados cadastrais do cliente e informações principais do orçamento.
                                </p>
                            </div>

                            <div class="switch-item pb-3 mb-4 border-bottom">
                                <div class="">
                                    <label class="switch-label" for="flagClienteNovo">Cliente Novo</label>
                                </div>
                                <div class="form-check form-switch mb-0">
                                    <input class="form-check-input" type="checkbox" id="flagClienteNovo"/>
                                </div>
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

                                <div class="col-12 col-lg-3 col-md-4">
                                    <label for="codCliente" class="form-label">Cód. Cliente <span class="required-field me-2">*</span></label>
                                    <div class="input-group">
                                        <button type="button" class="btn btn-sm btn-secondary" id="btnBuscarCliente" onclick="abrirModalCliente()">
                                            Buscar
                                            <i class="bi bi-box-arrow-up-right"></i>
                                        </button>
                                        <input required onblur="obterCliente()" type="text" class="form-control form-control-sm" id="codCliente" name="codCliente" required>
                                    </div>
                                    <small class="text-muted">Clique no botão para buscar ou digite um cód.</small>
                                </div>
                                
                                <div class="col-12 col-lg-3 col-md-4">
                                    <label for="codLoja" class="form-label">Cód. Loja <span class="required-field">*</span></label>
                                    <input type="text" onblur="obterCliente()" class="form-control form-control-sm" id="codLoja" name="codLoja">
                                </div>
                                
                                <div class="col-12 col-lg-3 col-md-4">
                                    <label for="nomeCliente" class="form-label">Nome do Cliente <span class="required-field">*</span></label>
                                    <input disabled type="text" class="form-control form-control-sm" id="nomeCliente" name="nomeCliente" disabled>
                                </div>

                                <div class="col-12 col-lg-3 col-md-4">
                                    <label for="CNPJ" class="form-label">CNPJ <span class="required-field">*</span></label>
                                    <input type="text" 
                                        class="form-control form-control-sm" 
                                        id="CNPJ" 
                                        name="CNPJ"
                                        placeholder="00.000.000/0000-00"
                                        maxlength="15"
                                        disabled>
                                </div>
                                
                                <div class="col-12 col-lg-3 col-md-4">
                                    <label for="codVendedor" class="form-label">Cód. Vendedor <span class="required-field me-2">*</span></label>
                                    <div class="input-group">
                                        <button type="button" class="btn btn-sm btn-secondary" id="btnBuscarVendedor" onclick="abrirModalVendedor()" disabled>
                                            Buscar
                                            <i class="bi bi-box-arrow-up-right"></i>
                                        </button>
                                        <input required onblur="obterVendedor()" type="text" class="form-control form-control-sm" id="codVendedor" name="codVendedor" disabled>
                                    </div>
                                    <small class="text-muted">Clique no botão para buscar ou digite um cód.</small>
                                </div>

                                <div class="col-12 col-lg-3 col-md-4">
                                    <label for="nomeVendedor" class="form-label">Nome do Vendedor <span class="required-field">*</span></label>
                                    <input type="text" class="form-control form-control-sm" id="nomeVendedor" name="nomeVendedor" disabled>
                                </div>

                                <div class="col-12 col-lg-3 col-md-6">
                                    <label for="telefone" class="form-label">Telefone </label>
                                    <input type="text" 
                                        class="form-control form-control-sm" 
                                        id="telefone" 
                                        name="telefone"
                                        placeholder="(00) 00000-0000"
                                        maxlength="15">
                                </div>

                                <div class="col-12 col-lg-3 col-md-6">
                                    <label for="contato" class="form-label">Contato </label>
                                    <input type="contato" 
                                        class="form-control form-control-sm" 
                                        id="contato" 
                                        name="contato">
                                </div>

                                <div class="col-12 col-lg-6 col-md-6">
                                    <label for="email" class="form-label">E-mail </label>
                                    <input disabled
                                        type="email" 
                                        class="form-control form-control-sm" 
                                        id="email" 
                                        name="email"
                                        placeholder="email@exemplo.com.br">
                                </div>

                                <div class="col-12 col-lg-6 col-md-6">
                                    <label for="CEP" class="form-label">CEP <span class="required-field">*</span></label>
                                    <input disabled
                                        type="text" 
                                        class="form-control form-control-sm" 
                                        id="CEP" 
                                        name="CEP" 
                                        placeholder="00000-000"
                                        data-mask="00000-000"
                                        maxlength="9">
                                </div>

                                <div class="col-12 col-lg-6 col-md-6">
                                    <label for="endereco" class="form-label">Endereço <span class="required-field">*</span></label>
                                    <input 
                                        disabled
                                        type="text" 
                                        class="form-control form-control-sm" 
                                        id="endereco" 
                                        name="endereco"
                                        placeholder="Rua">
                                </div>

                                <div class="col-12 col-md-2">
                                    <label for="numero" class="form-label">Número </label>
                                    <input 
                                        disabled
                                        type="text" 
                                        class="form-control form-control-sm" 
                                        id="numero" 
                                        name="numero"
                                        placeholder="Número">
                                </div>

                                <div class="col-12 col-md-4">
                                    <label for="complemento" class="form-label">Complemento </label>
                                    <input type="text" 
                                        class="form-control form-control-sm" 
                                        id="complemento" 
                                        name="complemento"
                                        placeholder="Complemento">
                                </div>

                                <div class="col-12 col-lg-4 col-md-8">
                                    <label for="bairro" class="form-label">Bairro <span class="required-field">*</span></label>
                                    <input disabled
                                        type="text" 
                                        class="form-control form-control-sm" 
                                        id="bairro" 
                                        name="bairro"
                                        placeholder="Bairro">
                                </div>

                                <div class="col-12 col-lg-3 col-md-4">
                                    <label for="UF" class="form-label">UF <span class="required-field">*</span></label>
                                    <select disabled class="form-select form-select-sm select2" id="UF" name="UF">
                                        <option value="">Selecione um estado</option>
                                        <option value="AC">Acre</option>
                                        <option value="AL">Alagoas</option>
                                        <option value="AP">Amapá</option>
                                        <option value="AM">Amazonas</option>
                                        <option value="BA">Bahia</option>
                                        <option value="CE">Ceará</option>
                                        <option value="DF">Distrito Federal</option>
                                        <option value="ES">Espírito Santo</option>
                                        <option value="GO">Goías</option>
                                        <option value="MA">Maranhão</option>
                                        <option value="MT">Mato Grosso</option>
                                        <option value="MS">Mato Grosso do Sul</option>
                                        <option value="MG">Minas Gerais</option>
                                        <option value="PA">Pará</option>
                                        <option value="PB">Paraíba</option>
                                        <option value="PR">Paraná</option>
                                        <option value="PE">Pernambuco</option>
                                        <option value="PI">Piauí</option>
                                        <option value="RJ">Rio de Janeiro</option>
                                        <option value="RN">Rio Grande do Norte</option>
                                        <option value="RS">Rio Grande do Sul</option>
                                        <option value="RO">Rondônia</option>
                                        <option value="RR">Roraíma</option>
                                        <option value="SC">Santa Catarina</option>
                                        <option value="SP">São Paulo</option>
                                        <option value="SE">Sergipe</option>
                                        <option value="TO">Tocantins</option>
                                    </select>
                                </div>

                                <div class="col-12 col-lg-5 col-md-12">
                                    <label for="municipioId" class="form-label">Município <span class="required-field">*</span></label>
                                    <select disabled class="form-select form-select-sm select2" id="municipioId" name="municipioId">
                                        <option value="">Selecione um Município</option>
                                    </select>
                                </div>
                            </div>
                        </div>
                
                        <!-- Aba Impostos e Condições -->
                        <div class="tab-content" data-tab-content="2">
                            <div class="section-header">
                                <div class="section-title">Impostos e condições</div>
                                <p class="section-subtitle">
                                    Condições comerciais, frete, pagamento e incidências de impostos.
                                </p>
                            </div>

                            <div class="row g-3">
                                <div class="col-lg-4 col-md-4">
                                    <label for="dataValidadeDe" class="form-label">Data Validade De <span class="required-field">*</span></label>
                                    <input type="date" class="form-control form-control-sm" id="dataValidadeDe" name="dataValidadeDe">
                                </div>
                                <div class="col-lg-4 col-md-4">
                                    <label for="dataValidadeAte" class="form-label">Data Validade Até <span class="required-field">*</span></label>
                                    <input type="date" class="form-control form-control-sm" id="dataValidadeAte" name="dataValidadeAte">
                                </div>
                                <div class="col-lg-4 col-md-4">
                                    <label for="tipoFrete" class="form-label">Tipo de Frete <span class="required-field">*</span></label>
                                    <select class="form-select form-select-sm select2" id="tipoFrete" name="tipoFrete">
                                        <option value="">Selecione o Tipo de Frete</option>
                                        <option value="CIF">CIF</option>
                                        <option value="FOB">FOB</option>
                                    </select>
                                </div>
                                <div class="col-lg-4 col-md-4">
                                    <label for="tipoCaminhao" class="form-label">Tipo de Caminhão <span class="required-field">*</span></label>
                                    <select class="form-select form-select-sm select2" id="tipoCaminhao" name="tipoCaminhao">
                                        <option value="">Selecione o Tipo de Caminhão</option>
                                    </select>
                                </div>
                                <div class="col-lg-4 col-md-4">
                                    <label for="condicaoPagamentoId" class="form-label">Condições de Pagamento <span class="required-field">*</span></label>
                                    <select class="form-select form-select-sm select2" id="condicaoPagamentoId" name="condicaoPagamentoId">
                                        <option value="">Condição de Pagamento</option>
                                    </select>
                                </div>
                                <div class="col-lg-4 col-md-4">
                                    <label for="temICMS" class="form-label">ICMS <span class="required-field">*</span></label>
                                    <select class="form-select form-select-sm select2" id="temICMS" name="temICMS">
                                        <option value="1">Sim</option>
                                        <option value="0" selected>Não</option>
                                    </select>
                                </div>
                                <div class="col-lg-4 col-md-6">
                                    <label for="aliquotaICMS" class="form-label">Alíquota de ICMS <span class="required-field">*</span></label>
                                    <input readonly
                                        type="number" 
                                        class="form-control form-control-sm" 
                                        id="aliquotaICMS" 
                                        name="aliquotaICMS"
                                        value="0"
                                        placeholder="Alíquota">
                                </div>
                                <div class="col-lg-4 col-md-4">
                                    <label for="temPIS" class="form-label">PIS <span class="required-field">*</span></label>
                                    <select class="form-select form-select-sm select2" id="temPIS" name="temPIS">
                                        <option value="1">Sim</option>
                                        <option value="0" selected>Não</option>
                                    </select>
                                </div>
                                <div class="col-lg-4 col-md-4">
                                    <label for="temCOFINS" class="form-label">COFINS <span class="required-field">*</span></label>
                                    <select class="form-select form-select-sm select2" id="temCOFINS" name="temCOFINS">
                                        <option value="1">Sim</option>
                                        <option value="0" selected>Não</option>
                                    </select>
                                </div>
                            </div>
                        </div>

                        <!-- Aba Itens -->
                        <div class="tab-content" data-tab-content="3">
                            <div class="section-header d-flex justify-content-between align-items-start flex-wrap gap-2">
                                <div>
                                    <div class="d-flex align-items-center gap-2">
                                        <div class="section-title">Produtos do orçamento</div>
                                        <span class="badge bg-secondary" id="contadorProdutos">0</span>
                                    </div>
                                    <p class="section-subtitle">
                                        Lista de produtos vinculados ao orçamento.
                                    </p>
                                </div>
                                <div class="d-flex gap-2">
                                    <button type="button" class="d-none btn btn-sm btn-outline-secondary" id="btnUltimoPedidoFT">
                                        <i class="bi bi-clock-history me-1"></i>
                                        <span>Último Pedido FT</span>
                                    </button>
                                    <button type="button" class="d-none btn btn-sm btn-primary" id="btnCalcular" onclick="abrirModalCalculo()">
                                        <i class="bi bi-calculator-fill me-1"></i>Calcular
                                    </button>
                                    <button type="button" class="d-none btn btn-sm btn-primary" id="btnAdicionarProduto" onclick="abrirModalProduto()">
                                        <i class="bi bi-plus-lg me-1"></i>Adicionar
                                    </button>
                                </div>
                            </div>

                            <div id="produtoListaContainer">
                                <!-- Lista populada dinamicamente -->
                            </div>
                        </div>

                        <!-- Aba Dados de Mensagens -->
                        <div class="tab-content" data-tab-content="4">
                            <div class="section-header d-flex justify-content-between align-items-start flex-wrap gap-2">
                                <div>
                                    <div class="d-flex align-items-center gap-2">
                                        <div class="section-title">Mensagens do orçamento</div>
                                        <span class="badge bg-secondary" id="contadorMensagens">0</span>
                                    </div>
                                    <p class="section-subtitle">
                                        Acompanhamento das mensagens registradas para o orçamento.
                                    </p>
                                </div>
                                <div class="d-flex gap-2">
                                    <button type="button" class="d-none btn btn-sm btn-primary" id="btnAdicionarMensagem" onclick="abrirModalMensagem()">
                                        <i class="bi bi-plus-lg me-1"></i>Adicionar
                                    </button>
                                </div>
                            </div>

                            <div id="mensagemListaContainer">
                                <!-- Lista populada dinamicamente -->
                            </div>
                        </div>
                    </form>

                    <div class="nav-footer">
                        <button type="button" class="btn btn-sm btn-outline-secondary" onclick="irAnterior()">
                            <i class="bi bi-arrow-left"></i> Voltar
                        </button>
                        <button type="button" class="btn btn-sm btn-primary" onclick="irProximo()">
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
    
    <!-- Modal de Vendedores -->
    <div class="modal fade" id="modalVendedor"  aria-labelledby="modalVendedorLabel" aria-hidden="true">
        <div class="modal-dialog modal-xl modal-dialog-centered modal-dialog-scrollable">
            <div class="modal-content">
                <div class="modal-header">
                    <h5 class="modal-title" id="modalVendedorLabel">
                        <i class="bi bi-search me-2"></i>Buscar Vendedor
                    </h5>
                    <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Fechar"></button>
                </div>
                <div class="modal-body">
                    <div class="mb-3">
                        <label for="vendedorBusca" class="form-label">Pesquisar Vendedor</label>
                        <div class="input-group">
                            <span class="input-group-text"><i class="bi bi-search"></i></span>
                            <input type="text" 
                                   class="form-control form-control-sm" 
                                   id="vendedorBusca" 
                                   placeholder="Digite o código ou nome para buscar..."
                                   autocomplete="off">
                        </div>
                    </div>

                    <div id="vendedorResultados" style="max-height: 400px; overflow-y: auto;">
                    </div>
                </div>
                <div class="modal-footer">
                    <button type="button" class="btn btn-sm btn-outline-secondary" data-bs-dismiss="modal">
                        Cancelar
                    </button>
                    <button type="button" class="btn btn-sm btn-primary" id="btnConfirmarVendedor" onclick="confirmarVendedor()" disabled>
                        <i class="bi bi-check-lg me-1"></i>Selecionar
                    </button>
                </div>
            </div>
        </div>
    </div>
    
    <!-- Modal de Produto -->
    <div class="modal fade" id="modalProduto"  aria-labelledby="modalProdutoLabel" aria-hidden="true">
        <div class="modal-dialog modal-xl">
            <div class="modal-content">
                <div class="modal-header">
                    <h5 class="modal-title" id="modalProdutoTitulo">
                        Adicionar Produto
                    </h5>
                    <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Fechar"></button>
                </div>

                <div class="modal-body" style=" overflow: auto; max-height: 70vh;">
                    <form id="formularioProduto" name="formularioProduto" novalidate>
                        <input type="hidden" name="hidProdutoIdOrcamento" id="hidProdutoIdOrcamento" value="" />
                        <input type="hidden" name="hidProdutoQuantidadeMinima" id="hidProdutoQuantidadeMinima" value="0" />
                        <input type="hidden" name="hidProdutoPrecoNet" id="hidProdutoPrecoNet" value="0" />
                        <input type="hidden" name="hidPrecoUnitario" id="hidPrecoUnitario" value="0" />
                        <input type="hidden" name="hidPrecoUnitarioNet" id="hidPrecoUnitarioNet" value="0" />
                        <input type="hidden" name="hidProdutoPrecoKGRefile" id="hidProdutoPrecoKGRefile" value="0" />
                        <input type="hidden" name="hidProdutoPrecoKGNetRefile" id="hidProdutoPrecoKGNetRefile" value="0" />
                        <input type="hidden" name="hidProdutoPrecoKGNet" id="hidProdutoPrecoKGNet" value="0" />
                        <input type="hidden" name="hidProdutoVolumeM2" id="hidProdutoVolumeM2" value="0" />
                        <input type="hidden" name="hidProdutoAreaBlank" id="hidProdutoAreaBlank" value="0" />
                        <input type="hidden" name="hidProdutoPesoBlank" id="hidProdutoPesoBlank" value="0" />
                        <input type="hidden" name="hidProdutoAliquotaIPI" id="hidProdutoAliquotaIPI" value="0" />
                        <input type="hidden" name="hidProdutoTipo" id="hidProdutoTipo" value="0" />
                        <input type="hidden" name="hidProdutoStatusCalculo" id="hidProdutoStatusCalculo" value="" />

                        <div class="switch-item pb-3 mb-4 border-bottom">
                            <div class="">
                                <label class="switch-label" for="flagProdutoNovo">Produto Novo</label>
                            </div>
                            <div class="form-check form-switch mb-0">
                                <input class="form-check-input" type="checkbox" id="flagProdutoNovo"/>
                            </div>
                        </div>

                        <div class="row g-3">
                            <div class="col-12 col-md-4">
                                <label for="produtoTipoItem" class="form-label">Tipo do item <span class="required-field">*</span></label>
                                <select class="form-select form-select-sm select2" name="produtoTipoItem" id="produtoTipoItem">
                                    <option value="CX" selected>Caixa</option>
                                    <option value="CH">Chapa</option>
                                    <option value="CJ">Conjunto</option>
                                    <option value="AC">Acessório</option>
                                </select>
                            </div>
                            <div class="col-12 col-md-4">
                                <label for="produtoTipoCaixa" class="form-label">Tipo de Caixa <span class="required-field">*</span></label>
                                <select class="form-select form-select-sm select2" name="produtoTipoCaixa" id="produtoTipoCaixa">
                                    <option value="1" selected>Maleta - Fundo (0200)</option>
                                    <option value="2">Maleta - Caixa Normal</option>
                                    <option value="4">Maleta - Abas Total (0203)</option>
                                    <option value="5">Corte e Vinco</option>
                                </select>
                            </div>
                            <div class="col-12 col-md-4">
                                <label for="produtoCod" class="form-label">Cód. Produto <span class="required-field me-2">*</span></label>
                                <div class="input-group">
                                    <button type="button" id="btnBuscarProduto" class="btn btn-sm btn-secondary" onclick="abrirModalPesquisaProduto()">
                                        Buscar 
                                        <i class="bi bi-box-arrow-up-right"></i>
                                    </button>
                                    <input required onblur="obterProduto()" type="text" class="form-control form-control-sm" id="produtoCod" name="produtoCod" required>
                                </div>
                                <small class="text-muted">Clique no botão para buscar ou digite um cód.</small>
                            </div>
                            <div class="col-12 col-md-4">
                                <label for="produtoEstrutura" class="form-label">Estrutura <span class="required-field">*</span></label>
                                <select class="form-select form-select-sm select2" name="produtoEstrutura" id="produtoEstrutura">
                                    <option value="">Selecione a estrutura</option>
                                </select>
                            </div>
                            <div class="col-12 col-md-4">
                                <label for="produtoReferencia" class="form-label">Referência <span class="required-field">*</span></label>
                                <input type="text" class="form-control form-control-sm" id="produtoReferencia" name="produtoReferencia">
                            </div>
                            <div class="col-12 col-md-4">
                                <label for="produtoQuantidade" class="form-label">Quantidade <span class="required-field">*</span></label>
                                <input value="0" type="number" class="form-control form-control-sm" id="produtoQuantidade" name="produtoQuantidade">
                            </div>
                            <div class="col-12 col-md-4">
                                <label for="produtoTipoPreco" class="form-label">Tipo de Preço <span class="required-field">*</span></label>
                                <select class="form-select form-select-sm select2" name="produtoTipoPreco" id="produtoTipoPreco">
                                    <option value="1" selected>Preço Unitário</option>
                                    <option value="2">Preço KG</option>
                                </select>
                            </div>
                            <div class="col-12 col-md-4">
                                <label for="produtoValorUnitario" id="produtoValorUnitarioLabel" class="form-label">Valor Unitário <span class="required-field">*</span></label>
                                <input value="0" type="number" class="form-control form-control-sm" id="produtoValorUnitario" name="produtoValorUnitario">
                            </div>
                            <div class="col-12 col-md-4">
                                <label for="produtoGramatura" class="form-label">Gramatura <span class="required-field">*</span></label>
                                <input value="0" disabled type="number" class="form-control form-control-sm" id="produtoGramatura" name="produtoGramatura">
                            </div>
                             <div class="col-12 col-md-4">
                                <label for="produtoComprimentoInterno" class="form-label">Comprimento Interno <span class="required-field">*</span></label>
                                <input value="0" type="number" class="form-control form-control-sm" id="produtoComprimentoInterno" name="produtoComprimentoInterno">
                            </div>
                             <div class="col-12 col-md-4">
                                <label for="produtoLarguraInterna" class="form-label">Largura Interna <span class="required-field">*</span></label>
                                <input value="0" type="number" class="form-control form-control-sm" id="produtoLarguraInterna" name="produtoLarguraInterna">
                            </div>
                             <div class="col-12 col-md-4">
                                <label for="produtoAlturaInterna" class="form-label">Altura Interna <span class="required-field">*</span></label>
                                <input value="0" type="number" class="form-control form-control-sm" id="produtoAlturaInterna" name="produtoAlturaInterna">
                            </div>
                            <div class="col-12 col-md-4">
                                <label for="produtoComprimentoChapa" class="form-label">Comprimento Chapa <span class="required-field">*</span></label>
                                <input value="0" type="number" class="form-control form-control-sm" id="produtoComprimentoChapa" name="produtoComprimentoChapa">
                            </div>
                            <div class="col-12 col-md-4">
                                <label for="produtoLarguraChapa" class="form-label">Largura Chapa <span class="required-field">*</span></label>
                                <input value="0" type="number" class="form-control form-control-sm" id="produtoLarguraChapa" name="produtoLarguraChapa">
                            </div>
                            <div class="col-12 col-md-4">
                                <label for="produtoArranjo" class="form-label">Arranjo <span class="required-field">*</span></label>
                                <input value="1" type="number" class="form-control form-control-sm" id="produtoArranjo" name="produtoArranjo">
                            </div>
                             <div class="col-12 col-md-4">
                                <label for="produtoPecasPorConjunto" class="form-label">Peças por Conjunto <span class="required-field">*</span></label>
                                <input value="1" type="number" class="form-control form-control-sm" id="produtoPecasPorConjunto" name="produtoPecasPorConjunto">
                            </div>
                            <div class="col-12 col-md-4">
                                <label for="produtoAreaRefile" class="form-label">Área Refile <span class="required-field">*</span></label>
                                <input value="0" disabled type="number" class="form-control form-control-sm" id="produtoAreaRefile" name="produtoAreaRefile">
                            </div>
                            <div class="col-12 col-md-4">
                                <label for="produtoPesoRefile" class="form-label">Peso Refile <span class="required-field">*</span></label>
                                <input value="0" disabled type="number" class="form-control form-control-sm" id="produtoPesoRefile" name="produtoPesoRefile">
                            </div>
                            <div class="col-lg-4 col-md-4">
                                <label for="produtoTemIPI" class="form-label">IPI <span class="required-field">*</span></label>
                                <select class="form-select form-select-sm select2" id="produtoTemIPI" name="produtoTemIPI">
                                    <option value="1">Sim</option>
                                    <option value="0" selected>Não</option>
                                </select>
                            </div>
                            <div class="col-12 col-md-4">
                                <label for="produtoAliquotaIPI" class="form-label">Alíquota IPI <span class="required-field">*</span></label>
                                <input value="0" disabled type="number" class="form-control form-control-sm" id="produtoAliquotaIPI" name="produtoAliquotaIPI">
                            </div>
                            <div class="col-12 col-md-4">
                                <label for="produtoQuantidadeCores" class="form-label">Quantidade de Cores</label>
                                <input value="0" type="number" class="form-control form-control-sm" id="produtoQuantidadeCores" name="produtoQuantidadeCores">
                            </div>
                            <div class="col-12 col-md-4">
                                <label for="produtoPedidoCompra" class="form-label">Pedido de Compra </label>
                                <input type="number" class="form-control form-control-sm" id="produtoPedidoCompra" name="produtoPedidoCompra">
                            </div>
                            <div class="col-12 col-md-4">
                                <label for="produtoItemPedido" class="form-label">Item do Pedido </label>
                                <input type="number" class="form-control form-control-sm" id="produtoItemPedido" name="produtoItemPedido">
                            </div>
                            <div class="col-12">
                                <label for="produtoObservacao" class="form-label">Observação</label>
                                <textarea class="form-control form-control-sm" id="produtoObservacao" name="produtoObservacao" rows="4"></textarea>
                            </div>
                        </div>
                    </form>
                </div>
                <div class="modal-footer">
                    <button type="button" class="btn btn-sm btn-outline-secondary" data-bs-dismiss="modal">
                        Cancelar
                    </button>
                    <button type="button" class="btn btn-sm btn-primary" id="btnSalvarProduto">
                        <i class="bi bi-plus-lg me-1"></i>Adicionar
                    </button>
                </div>
            </div>
        </div>
    </div>

    <!-- Modal de Mensagem -->
    <div class="modal fade" id="modalMensagem"  aria-labelledby="modalMensagemLabel" aria-hidden="true">
        <div class="modal-dialog modal-xl">
            <div class="modal-content">
                <div class="modal-header">
                    <h5 class="modal-title" id="modalMensagemTitulo">
                       Adicionar Mensagem
                    </h5>
                    <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Fechar"></button>
                </div>
                <div class="modal-body" style=" overflow: auto; max-height: 70vh;">
                    <form id="formularioMensagem" name="formularioMensagem" novalidate>
                        <div class="row g-3">
                            <div class="col-12">
                                <label for="mensagemNivel" class="form-label">Para quem deseja enviar <span class="required-field">*</span> </label>
                                <select class="form-select form-select-sm select2" name="mensagemNivel" id="mensagemNivel">
                                    <option value="0">Todos</option>
                                    <option value="1">Analistas e Aprovadores</option>
                                </select>
                            </div>
                            <div class="col-12">
                                <label for="mensagemDescricao" class="form-label">Descrição</label>
                                <textarea class="form-control form-control-sm" id="mensagemDescricao" name="mensagemDescricao" rows="4"></textarea>
                            </div>
                        </div>
                    </form>
                </div>
                <div class="modal-footer">
                    <button type="button" class="btn btn-sm btn-outline-secondary" data-bs-dismiss="modal">
                        Cancelar
                    </button>
                    <button type="button" class="btn btn-sm btn-primary" id="btnSalvarMensagem">
                        <i class="bi bi-plus-lg me-1"></i>Adicionar
                    </button>
                </div>
            </div>
        </div>
    </div>

    <!-- Modal de Cálculo -->
    <div class="modal fade" id="modalCalculo"  aria-labelledby="modalCalculoLabel" aria-hidden="true">
        <div class="modal-dialog modal-lg">
            <div class="modal-content">
                <div class="modal-header">
                    <h5 class="modal-title" id="modalCalculoTitulo">
                        <i class="bi bi-calculator me-2"></i>Realizar Cálculo
                    </h5>
                    <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Fechar"></button>
                </div>
                <div class="modal-body" style=" overflow: auto; max-height: 70vh;">
                    <form id="formularioCalculo" name="formularioCalculo" novalidate>
                        <div class="row g-3">
                            <div class="col-12">
                                <label for="calculoTipo" class="form-label">Tipo de Cálculo</label>
                                <select class="form-select form-select-sm select2" name="calculoTipo" id="calculoTipo">
                                    <option selected value="MC">Margem Contribuição</option>
                                    <option value="PK">Preço KG</option>
                                    <option value="VU">Valor Unitário</option>
                                </select>
                            </div>
                            <!-- Cálculo de Margem é obrigatório -->
                            <div class="col-12">
                                <label for="calculoValorBase" class="form-label">Valor Base</label>
                                <input 
                                    type="text" 
                                    class="form-control form-control-sm" 
                                    id="calculoValorBase" 
                                    name="calculoValorBase"
                                    placeholder="Número">
                            </div>
                            <div class="form-text mt-3" style="font-size:.85rem;">
                                <i class="bi bi-info-circle me-1"></i> Caso nenhum valor seja preenchido será usado o valor padrão do cadastro
                            </div>
                        </div>
                    </form>
                </div>
                <div class="modal-footer">
                    <button type="button" class="btn btn-sm btn-outline-secondary" data-bs-dismiss="modal">
                        Cancelar
                    </button>
                    <button type="button" class="btn btn-sm btn-primary" id="btnConfirmarCalculo">
                        <i class="bi bi-plus-check me-1"></i>Calcular
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
    
    <!-- Modal de Resultados do Processamento de calculo -->
    <div class="modal fade" id="modalResultadosCalculo" data-backdrop="static" data-keyboard="false"  role="dialog" aria-labelledby="modalResultadosCalculoLabel" aria-hidden="true">
        <div class="modal-dialog modal-xl modal-dialog-centered modal-dialog-scrollable">
            <div class="modal-content">
                <div class="modal-header">
                    <h5 class="modal-title" id="modalResultadosCalculoLabel">
                        Cálculos
                    </h5>
                    <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Close"></button>
                </div>
                
                <div id="resultadosCalculo" class="modal-body p-0"> </div>
            </div>
        </div>
    </div>

    <!-- Modal de Último Pedido FT -->
    <div class="modal fade" id="modalUltimoPedidoFT"  aria-labelledby="modalUltimoPedidoFTLabel" aria-hidden="true">
        <div class="modal-dialog modal-xl">
            <div class="modal-content">
                <div class="modal-header">
                    <h5 class="modal-title">
                        Último Pedido FT
                    </h5>
                    <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Fechar"></button>
                </div>
                <div class="modal-body">
                    <table class="table table-hover mb-0" id="gradeUltimoPedidoFT">
                        <thead>
                            <tr>
                                <th>Produto</th>
                                <th>Referência</th>
                                <th>Nº Último Pedido</th>
                                <th>Preço Unitário</th>
                                <th>Condição Pagamento</th>
                                <th>Preço KG Net</th>
                                <th>Data de Emissão</th>
                            </tr>
                        </thead>
                        <tbody>
                            <tr>
                                <td colspan="28" class="text-center text-muted py-4">
                                    Nenhum registro encontrado
                                </td>
                            </tr>
                        </tbody>
                    </table>
                </div>
            </div>
        </div>
    </div>
    

    <!-- Modal de Detalhes da Mensagem -->
    <div class="modal fade" id="modalDetalheMensagem" tabindex="-1" aria-labelledby="modalDetalheMensagemTitulo" aria-hidden="true">
        <div class="modal-dialog modal-lg modal-dialog-scrollable">
            <div class="modal-content">
                <div class="modal-header">
                    <h5 class="modal-title" id="modalDetalheMensagemTitulo">
                        Detalhes da Mensagem
                    </h5>
                    <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Fechar"></button>
                </div>
                <div class="modal-body">
                    <div id="modalDetalheMensagemConteudo" class="detalhe-conteudo"></div>
                </div>
                <div class="modal-footer">
                    <button type="button" class="btn btn-sm btn-secondary" data-bs-dismiss="modal">Fechar</button>
                </div>
            </div>
        </div>
    </div>

</body>
</html>