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
            font-size: .8rem;
        }

        .container {
            max-width: 90%;
        }

        .table th,
        .table td {
            white-space: nowrap;
            vertical-align: top;
        }

        thead {
            vertical-align: middle !important;
            position: sticky; 
            inset-block-start: 0; 
            background: #f5f5f5;
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


        .form-label {
            font-weight: 600;
        }

        .required-field {
            color: #dc3545;
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


         /* Layout */
        .form-container {
            background: #fff;
            border-radius: 8px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
            overflow: hidden;
        }
        .form-layout { display: grid; grid-template-columns: 270px minmax(0, 1fr); }
        .tabs-container { background-color: #f8f9fa; border-right: 1px solid #e9ecef; overflow-y: auto; }
        .tabs { display: flex; flex-direction: column; gap: 3px; padding: 14px 10px; }
        .tab-group-title { color: #9aa3af; font-size: 11px; font-weight: 600; letter-spacing: .06em; margin: 14px 10px 6px; text-transform: uppercase; }
        .tab-group-title:first-child { margin-top: 0; }
        .tab-button { align-items: center; background: transparent; border: none; border-radius: 7px; color: #495057; cursor: pointer; display: flex; font-weight: 500; gap: 9px; justify-content: flex-start; padding: 9px 11px; text-align: left; transition: all .2s ease; white-space: nowrap; width: 100%; }
        .tab-button:hover { background-color: #eef2f7; }
        .tab-button.active { background-color: #e7f1ff; color: #0d6efd; font-weight: 700; }
        .tab-button.completed { color: #198754; }
        .tab-button.invalid { background-color: #fff5f5; color: #dc3545; }
        .tab-icon { flex-shrink: 0; font-size: 16px; width: 18px; }
        .tab-title { flex: 1; font-size: 14px; }
        .tab-meta { background: #e9ecef; border-radius: 999px; color: #6c757d; font-size: 10px; font-weight: 700; margin-left: auto; padding: 1px 7px; }
        .tab-button.active .tab-meta { background: #d7eaff; color: #0d6efd; }
        .tab-content-wrapper { background: #fff; min-width: 0; }
        .tab-content { animation: fadeIn .25s ease; display: none; padding: 30px; height: 80vh; overflow: auto; }
        .tab-content.active { display: block; }
        .nav-footer { position: sticky; bottom: 0; background: #fff; border-top: 1px solid #dee2e6; padding: 10px 20px; display: flex; justify-content: flex-end; gap: 10px; z-index: 10; }
        .section-header { border-bottom: 1px solid #e9ecef; margin-bottom: 22px; padding-bottom: 14px; }
        .section-title { color: #343a40; font-size: 1.05rem; font-weight: 700; margin-bottom: 3px; }
        .section-subtitle { color: #6c757d; font-size: .825rem; margin: 0; }
        .readonly-alert { background: #f8f9fa; border: 1px dashed #ced4da; border-radius: 8px; color: #6c757d; padding: 12px 14px; }
        .just-validate-error-label { color: #f87171 !important; display: block; font-size: .8rem; margin-top: 3px; }
        .just-validate-error-field { border-color: #f87171 !important; }
        @keyframes fadeIn { from { opacity: 0; transform: translateY(6px); } to { opacity: 1; transform: translateY(0); } }
        
        @media (max-width: 1400px) {
            .container { max-width: 95%; }
            .form-layout { grid-template-columns: 1fr; }
            .tabs-container { border-bottom: 1px solid #e9ecef; border-right: none; overflow-x: auto; }
            .tabs { flex-direction: row; min-width: max-content; padding: 8px; }
            .tab-group-title, .tab-meta { display: none; }
            .tab-button { justify-content: center; min-width: 135px; text-align: center; width: auto; }
            .tab-title { flex: initial; }
        }

        #gradeProdutos th:first-child,
        #gradeProdutos td:first-child {
            position: sticky;
            left: 0;
            z-index: 5;
            white-space: nowrap;
        }
        
        #gradeProdutos td:first-child {
            background: linear-gradient(to left, rgba(0,0,0,0.1) 0%, transparent 1px),#fff;
            background-origin: border-box;
            background-clip: padding-box;
        }
        
        #gradeProdutos thead th:first-child {
            background: linear-gradient(to left, rgba(0,0,0,0.1) 0%, transparent 1px), #f5f5f5;
            background-origin: border-box;
            background-clip: padding-box;
            z-index: 9999;
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

        thead {
            z-index: 10;
        }
    </style>

    <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.0.2/dist/js/bootstrap.bundle.min.js" integrity="sha384-MrcW6ZMFYlzcLA8Nl+NtUVF0sA7MsXsP1UyJoMp4YLEuNSfAP+JcXn/tWtIaxVXM" crossorigin="anonymous"></script>
    <script src="https://cdn.jsdelivr.net/npm/select2@4.1.0-rc.0/dist/js/select2.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/jquery.mask/1.14.15/jquery.mask.min.js"></script>
    <script src="https://unpkg.com/just-validate@latest/dist/just-validate.production.min.js"></script>

    <script src="/corporativo/components/upload/fileUpload.js?v=<%=Now()%>"></script>
    <script src="/corporativo/components/toast.js?v=<%=Now()%>"></script>
    <script src="/corporativo/components/confirmation.js?v=<%=Now()%>"></script>
    <script src="/corporativo/funcoes/funcoesJAVASCRIPT.js?v=<%=Now()%>"></script>

    <script>
        let numeroOrcamento = 0;
        let abaAtual = 1;
        let totalAbas = 4;
        let validacoes = {};

        let listaProdutos = [];
        let listaMensagens = [];
        
        // Auxiliar para comparação de preços
        let listaUltimoPedidoFT = []
        let flagAprovarAvaliacao = null;

        function anexarEventos() {
            $('.tab-button').each(function() {
                $(this).on('click', async function() {
                    const tabNumber = parseInt($(this).data('tab'));
                    navegarParaAba(tabNumber);
                });
            });

            $('#btnUltimoPedidoFT').on('click', function () {
                abrirModalUltimoPedidoFT();
            });

            $('#btnAprovar').on('click', function() { 
                abrirModalAvaliacao(true)
            });

            $('#btnRejeitar').on('click', function() { 
                abrirModalAvaliacao(false)
            });

            $('#btnConfirmarAvaliacao').on('click', function() { 
                avaliarOrcamento(flagAprovarAvaliacao)
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

        }
        
        function carregarDadosOrcamento() {
            const params = { 
                acao: 'obterPorNumeroOrcamento', 
                numeroOrcamento: numeroOrcamento, 
            }

            $.ajax({
                url: 'OrcamentoRepositorio.asp',
                type: 'POST',
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
                    console.error('Erro ao carregar dados para edição:', error);
                    Toast.error('Erro ao carregar dados para edição');
                }
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


        function avaliarOrcamento(flagAprovar) {
            if (!validarAcao(flagAprovar)) {
                return;
            }
            exibirLoading(true)

            $.ajax({
                url: 'OrcamentoRepositorio.asp',
                type: 'POST',
                dataType: 'json',
                data: {
                    numeroOrcamento: numeroOrcamento,
                    acao: 'avaliar',
                    flagAprovar: flagAprovar ? 1 : 0,
                    observacao: $('#avaliacaoObservacao').val() || '',
                    mensagens: JSON.stringify(listaMensagens),
                },
                success: (response) => {
                    if (!response.sucesso) {
                        exibirLoading(false)
                        Toast.error(response.mensagem);
                        return;
                    }
                         
                    setTimeout(() => {
                        navegarParaGrade();
                    }, 1000)
                },
                error: (xhr, status, error) => {
                    console.error('Erro:', error);
                    Toast.error('Erro ao processar validação: ' + error);
                },
                complete: () => {
                    exibirLoading(false)
                }
            });
        }

        function preencherFormulario(dados) {
            var p1 = carregarMunicipios(dados.codPlanta, dados.codMunicipio, dados.UF);
            var p2 = carregarCondicoesPagamento(dados.codPlanta);
            var p3 = carregarTiposCaminhao(dados.codPlanta);
            var p4 = carregarEstruturas(dados.codPlanta);
            var p5 = carregarUltimoPedidoFT(dados.codPlanta, dados.codCliente, dados.codLoja, dados.itens);

            $.when(p1, p2, p3, p4, p5).then(function() {
                $('#codCliente').val(dados.codCliente)
                $('#codLoja').val(dados.codLoja)
                $('#nomeCliente').val(dados.nomeCliente)
                $('#CNPJ').val(dados.CNPJ)
                $('#codVendedor').val(dados.codVendedor)
                $('#nomeVendedor').val(dados.nomeVendedor)
                $('#telefone').val(dados.telefone)
                $('#contato').val(dados.contato)
                $('#email').val(dados.email)
                $('#CEP').val(dados.CEP)
                $('#endereco').val(dados.endereco)
                $('#numero').val(dados.numero)
                $('#complemento').val(dados.complemento)
                $('#bairro').val(dados.bairro)
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
                }, 500);
            });
        }

        function renderizarProdutos() {
            const $container = $('#produtoListaContainer');
            const $contador = $('#contadorProdutos');
            
            $contador.text(listaProdutos.length);
            $('#menuContadorProdutos').text(listaProdutos.length);

            if (listaProdutos.length === 0) {
                $container.html(`
                    <div class="text-center py-5">
                        <p class="text-muted mt-3 mb-0">Nenhum produto adicionado.</p>
                    </div>
                `);
                return;
            }

            let html = `
                <div style="max-height: 60vh; overflow: auto;" class="table-responsive">
                    <table id="gradeProdutos" class="table table-hover table-sm mb-0">
                        <thead class="table-light">
                            <tr>
                                <th>Referência</th>
                                <th>Produto</th>
                                <th>Quantidade</th>
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
                                <th>Preço NET</th>
                                <th>Preço KG Refile</th>
                                <th>Preço KG Net Refile</th>
                                <th>Preço KG Net</th>
                                <th>Qtd Cores</th>
                                <th>Volume M2</th>
                                <th>Margem</th>
                                <th class="text-center produto-acoes">Ações</th>
                            </tr>
                        </thead>
                        <tbody>
            `;

            /* Auxiliar para comparação */
            function compararPrecoKGNet(precoAtual, precoAnterior) {
                if (precoAnterior === null || isNaN(precoAnterior)) {
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
            
            
            listaProdutos.forEach((produto, index) => {
                const precoAnterior = listaUltimoPedidoFT.find(i => i.produto === produto.codigo)?.precoKGNetBlank;
                const comparacao = compararPrecoKGNet(produto.precoKGNetBlank, precoAnterior);
                const textoTooltip = comparacao.classe === 'up' ? 'Valor maior que último pedido' : comparacao.classe === 'down' ? 'Valor menor que último pedido' : '';

                html += `
                    <tr>
                        ${criarCelulaDetalhe(produto.referencia, 'Referência')}
                        <td>${produto.codigo}</td> 
                        <td>${produto.quantidade || 0}</td> 
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
                            <div class="odd-box ${comparacao.classe}" data-bs-toggle="tooltip" data-bs-placement="top" title="${textoTooltip}">
                                <div>${produto.precoKGNetBlank || 0}</div>
                                <i class="${comparacao.icone}"></i>
                            </div>
                        </td>
                        <td>${produto.quantidadeCores || 0}</td>
                        <td>${produto.volumeM2 || 0}</td>
                        <td>${produto.margem || 0}</td>
                        <td class="text-center produto-acoes">
                            <button type="button" class="btn btn-sm btn-primary me-1" onclick="abrirModalResultadosCalculo(${produto.idOrcamento})" title="Visualizar Cálculos">
                                <i class="bi bi-eye"></i>
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

        function renderizarMensagens() {
            const $container = $('#mensagemListaContainer');
            const $contador = $('#contadorMensagens');
            
            $contador.text(listaMensagens.length);
            $('#menuContadorMensagens').text(listaMensagens.length);

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
                        <td>
                            <div class="truncate-text" title="${item.descricao || ''}">
                                ${item.descricao || '-'}
                            </div>
                        </td>
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

            listaMensagens.push(mensagem);
            Toast.success('Mensagem adicionada com sucesso!');

            renderizarMensagens();

            const modal = bootstrap.Modal.getInstance(document.getElementById('modalMensagem'));
            modal.hide();
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

        function abrirModalResultadosCalculo(idOrcamento = "") {
            const modal = new bootstrap.Modal(document.getElementById('modalResultadosCalculo'));
            modal.show();
            
            $("#resultadosCalculo").html('<div class="text-center p-4"> <div class="spinner-border text-primary mb-3" role="status"></div></div>');

            $.ajax({
                url: 'OrcamentoCalculoTrimboxResultados.asp?v=' + Math.random(),
                type: 'POST',
                data: { 
                    idOrcamentoSknet: idOrcamento, 
                },
                success: function(response) {
                    $('#resultadosCalculo').html(response);
                   
                },
                error: function(xhr, status, error) {
                    $('#resultadosCalculo').html('<div class="alert alert-danger">Erro ao carregar resultados: ' + error + '</div>');
                }
            });
        }

        function abrirModalDetalhe(botao) {
            const titulo = $(botao).attr('data-titulo') || 'Detalhes';
            const detalhe = $(botao).attr('data-detalhe') || '-';

            $('#modalDetalheTituloTexto').text(titulo);
            $('#modalDetalheConteudo').text(detalhe);

            const modal = new bootstrap.Modal(document.getElementById('modalDetalheTexto'));
            modal.show();
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

        function irProximo() { 
            navegarParaAba(abaAtual + 1); 
        }

        function irAnterior(){
            navegarParaAba(abaAtual - 1);
        }

        function abrirModalAvaliacao(flagAprovar) {
            flagAprovarAvaliacao = flagAprovar;
            $('#avaliacaoObservacao').val('');
            $('#avaliacaoObservacao').removeClass('just-validate-error-field');
            $('#avaliacaoObservacao').next('.just-validate-error-label').remove();

            if (flagAprovar) {
                $('#modalAvaliacaoTitulo').html('<i class="bi bi-check-lg me-2 text-success"></i>Aprovar orçamento');
                $('#modalAvaliacaoDescricao').text('Informe uma observação, se necessário, para registrar a aprovação do orçamento.');
                $('#avaliacaoObservacaoObrigatoriedade').text('(opcional para aprovação)');
                $('#btnConfirmarAvaliacao')
                    .removeClass('btn-danger')
                    .addClass('btn-success')
                    .html('<i class="bi bi-check-lg me-1"></i>Aprovar');
            } else {
                $('#modalAvaliacaoTitulo').html('<i class="bi bi-x-lg me-2 text-danger"></i>Rejeitar orçamento');
                $('#modalAvaliacaoDescricao').text('Informe a observação obrigatória para registrar a rejeição do orçamento.');
                $('#avaliacaoObservacaoObrigatoriedade').text('(obrigatória para rejeição)');
                $('#btnConfirmarAvaliacao')
                    .removeClass('btn-success')
                    .addClass('btn-danger')
                    .html('<i class="bi bi-x-lg me-1"></i>Rejeitar');
            }

            const modal = new bootstrap.Modal(document.getElementById('modalAvaliacao'));
            modal.show();
        }

        function validarAcao(flagAprovar) {
            const observacao = ($('#avaliacaoObservacao').val() || '').trim();
            $('#avaliacaoObservacao').removeClass('just-validate-error-field');
            $('#avaliacaoObservacao').next('.just-validate-error-label').remove();

            if (!flagAprovar && !observacao) {
                $('#avaliacaoObservacao')
                    .addClass('just-validate-error-field')
                    .after('<div class="just-validate-error-label">Observação obrigatória para rejeição</div>');
                Toast.warning('Observação obrigatória para rejeição');
                return false;
            }

            marcarAbaComoCompleta(1);
            return true;
        }

        function navegarParaGrade() {
            window.location.href = `OrcamentoGrade.asp`;
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
                         data-titulo="${tituloSeguro}"
                         data-detalhe="${textoSeguro}">
                        ${textoSeguro}
                    </div>
                </td>
            `;
        }

        function inicializarValidacoes() {   
            const naoVazio = (msg) => ({
                validator: (value) => value !== null && value !== undefined && String(value).trim() !== '',
                errorMessage: msg
            });

            validacoes.modalMensagem = new JustValidate('#formularioMensagem', {
                lockForm: false, 
            });

            validacoes.modalMensagem
                .addField('#mensagemNivel', [naoVazio('Para quem deseja enviar é obrigatório')])
                .addField('#mensagemDescricao', [naoVazio('Mensagem é obrigatória')]);
        }

        $(document).ready(function() {   
            const urlParams = new URLSearchParams(window.location.search);
            const paramNumeroOrcamento = urlParams.get('numeroOrcamento');

            if(!paramNumeroOrcamento) {
                Toast.error('Orçamento inválido')
                return;
            }

            numeroOrcamento = paramNumeroOrcamento;
            $('#titulo').text(`Avaliar Orçamento #${numeroOrcamento}`)

            inicializarValidacoes();
            $('.select2').select2({ theme: 'bootstrap-5', width: '100%', dropdownCssClass: 'z-index-9999' });
            navegarParaAba(1);
            carregarDadosOrcamento()
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
            <div>
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
                        <button type="button" class="tab-button active" data-tab="1"><i class="tab-icon bi bi-check2-square"></i><span class="tab-title">Avaliação</span></button>
                        <div class="tab-group-title">Consulta</div>
                        <button type="button" class="tab-button" data-tab="2"><i class="tab-icon bi bi-clipboard-check"></i><span class="tab-title">Cadastro</span></button>
                        <button type="button" class="tab-button" data-tab="3"><i class="tab-icon bi bi-currency-dollar"></i><span class="tab-title">Impostos e Condições</span></button>
                        <div class="tab-group-title">Histórico</div>
                        <button type="button" class="tab-button" data-tab="4"><i class="tab-icon bi bi-envelope"></i><span class="tab-title">Mensagens</span><span class="tab-meta" id="menuContadorMensagens">0</span></button>
                    </div>
                </div>


                <div class="tab-content-wrapper">
                    <div class="tab-content active" data-tab-content="1">
                        <div class="section-header d-flex justify-content-between align-items-start flex-wrap gap-2">
                            <div>
                                <div class="d-flex align-items-center gap-2">
                                    <div class="section-title">Avaliação do orçamento</div>
                                </div>
                                <p class="section-subtitle">
                                    <p class="section-subtitle">Registre a aprovação ou rejeição do orçamento.</p>
                                </p>
                            </div>
                            <div class="d-flex gap-2">
                                <button type="button" class="btn btn-sm btn-secondary" id="btnUltimoPedidoFT">
                                    <i class="bi bi-clock-history me-1"></i>
                                    <span>Último Pedido FT</span>
                                </button>
                            </div>
                        </div>

                        <div id="produtoListaContainer">
                            <!-- Lista populada dinamicamente -->
                        </div>

                    </div>

                    <!-- Aba Dados Cadastrais -->
                    <div class="tab-content" data-tab-content="2">
                        <div class="section-header">
                            <div class="section-title">Cadastro do orçamento</div>
                            <p class="section-subtitle">
                                Dados cadastrais do cliente e informações principais do orçamento.
                            </p>
                        </div>

                        <div class="row g-3">
                            <div class="col-12">
                                <label for="codPlanta" class="form-label">Planta </label>
                                <select class="form-select form-select-sm select2" name="codPlanta" id="codPlanta" disabled>
                                    <option value="">Selecione a planta</option>
                                    <option value="00">Pirapetinga</option>
                                    <option value="04">Uberaba</option>
                                    <option value="05">Bento</option>
                                    <option value="06">Maranguape</option>
                                    <option value="07">Saquarema</option>
                                </select>
                            </div>

                            <div class="col-12 col-lg-3 col-md-4">
                                <label for="codCliente" class="form-label">Cód. Cliente </label>
                                <input required type="text" class="form-control form-control-sm" id="codCliente" name="codCliente" disabled>
                            </div>
                            
                            <div class="col-12 col-lg-3 col-md-4">
                                <label for="codLoja" class="form-label">Cód. Loja</label>
                                <input type="text" class="form-control form-control-sm" id="codLoja" name="codLoja" disabled>
                            </div>
                            
                            <div class="col-12 col-lg-3 col-md-4">
                                <label for="nomeCliente" class="form-label">Nome do Cliente</label>
                                <input disabled type="text" class="form-control form-control-sm" id="nomeCliente" name="nomeCliente" disabled>
                            </div>

                            <div class="col-12 col-lg-3 col-md-4">
                                <label for="CNPJ" class="form-label">CNPJ</label>
                                <input type="text" 
                                    class="form-control form-control-sm" 
                                    id="CNPJ" 
                                    name="CNPJ"
                                    placeholder="00.000.000/0000-00"
                                    maxlength="15"
                                    disabled>
                            </div>
                                
                            <div class="col-12 col-lg-3 col-md-4">
                                <label for="codVendedor" class="form-label">Cód. Vendedor <span class="required-field">*</span></label>
                                <input type="text" class="form-control form-control-sm" id="codVendedor" name="codVendedor" disabled>
                            </div>

                            <div class="col-12 col-lg-3 col-md-4">
                                <label for="nomeVendedor" class="form-label">Nome do Vendedor <span class="required-field">*</span></label>
                                <input type="text" class="form-control form-control-sm" id="nomeVendedor" name="nomeVendedor" disabled>
                            </div>

                            <div class="col-12 col-lg-3 col-md-6">
                                <label for="telefone" class="form-label">Telefone <span class="required-field">*</span></label>
                                <input type="text" 
                                    class="form-control form-control-sm" 
                                    id="telefone" 
                                    name="telefone"
                                    placeholder="(00) 00000-0000"
                                    maxlength="15" disabled>
                            </div>

                            <div class="col-12 col-lg-3 col-md-6">
                                <label for="contato" class="form-label">Contato <span class="required-field">*</span></label>
                                <input type="contato" 
                                    class="form-control form-control-sm" 
                                    id="contato" 
                                    name="contato" disabled>
                            </div>

                            <div class="col-12 col-lg-6 col-md-6">
                                <label for="email" class="form-label">E-mail <span class="required-field">*</span></label>
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
                                <label for="numero" class="form-label">Número <span class="required-field">*</span></label>
                                <input 
                                    disabled
                                    type="text" 
                                    class="form-control form-control-sm" 
                                    id="numero" 
                                    name="numero"
                                    placeholder="Número">
                            </div>

                            <div class="col-12 col-md-4">
                                <label for="complemento" class="form-label">Complemento <span class="required-field">*</span></label>
                                <input type="text" 
                                    class="form-control form-control-sm" 
                                    id="complemento" 
                                    name="complemento"
                                    placeholder="Complemento" disabled>
                            </div>

                            <div class="col-12 col-lg-4 col-md-8">
                                <label for="bairro" class="form-label">Bairro <span class="required-field">*</span></label>
                                <input disabled
                                    type="text" 
                                    class="form-control form-control-sm" 
                                    id="bairro" 
                                    name="bairro"
                                    placeholder="Bairro" disabled>
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
                    <div class="tab-content" data-tab-content="3">
                        <div class="section-header">
                            <div class="section-title">Impostos e condições</div>
                            <p class="section-subtitle">
                                Condições comerciais, frete, pagamento e incidências de impostos.
                            </p>
                        </div>

                        <div class="row g-3">
                            <div class="col-lg-4 col-md-4">
                                <label for="dataValidadeDe" class="form-label">Data Validade De <span class="required-field">*</span></label>
                                <input disabled type="date" class="form-control form-control-sm" id="dataValidadeDe" name="dataValidadeDe">
                            </div>
                            <div class="col-lg-4 col-md-4">
                                <label for="dataValidadeAte" class="form-label">Data Validade Até <span class="required-field">*</span></label>
                                <input disabled type="date" class="form-control form-control-sm" id="dataValidadeAte" name="dataValidadeAte">
                            </div>
                            <div class="col-lg-4 col-md-4">
                                <label for="tipoFrete" class="form-label">Tipo de Frete <span class="required-field">*</span></label>
                                <select disabled class="form-select form-select-sm select2" id="tipoFrete" name="tipoFrete">
                                    <option value="">Selecione o Tipo de Frete</option>
                                    <option value="CIF">CIF</option>
                                    <option value="FOB">FOB</option>
                                </select>
                            </div>
                            <div class="col-lg-4 col-md-4">
                                <label for="tipoCaminhao" class="form-label">Tipo de Caminhão <span class="required-field">*</span></label>
                                <select disabled class="form-select form-select-sm select2" id="tipoCaminhao" name="tipoCaminhao">
                                    <option value="">Selecione o Tipo de Caminhão</option>
                                </select>
                            </div>
                            <div class="col-lg-4 col-md-4">
                                <label for="condicaoPagamentoId" class="form-label">Condições de Pagamento <span class="required-field">*</span></label>
                                <select disabled class="form-select form-select-sm select2" id="condicaoPagamentoId" name="condicaoPagamentoId">
                                    <option value="">Condição de Pagamento</option>
                                </select>
                            </div>
                            <div class="col-lg-4 col-md-4">
                                <label for="temICMS" class="form-label">ICMS <span class="required-field">*</span></label>
                                <select disabled class="form-select form-select-sm select2" id="temICMS" name="temICMS">
                                    <option value="1">Sim</option>
                                    <option value="0" selected>Não</option>
                                </select>
                            </div>
                            <div class="col-lg-4 col-md-6">
                                <label for="aliquotaICMS" class="form-label">Alíquota de ICMS <span class="required-field">*</span></label>
                                <input disabled
                                    type="number" 
                                    class="form-control form-control-sm" 
                                    id="aliquotaICMS" 
                                    name="aliquotaICMS"
                                    value="0"
                                    placeholder="Alíquota">
                            </div>
                            <div class="col-lg-4 col-md-4">
                                <label for="temPIS" class="form-label">PIS <span class="required-field">*</span></label>
                                <select disabled class="form-select form-select-sm select2" id="temPIS" name="temPIS">
                                    <option value="1">Sim</option>
                                    <option value="0" selected>Não</option>
                                </select>
                            </div>
                            <div class="col-lg-4 col-md-4">
                                <label for="temCOFINS" class="form-label">COFINS <span class="required-field">*</span></label>
                                <select disabled class="form-select form-select-sm select2" id="temCOFINS" name="temCOFINS">
                                    <option value="1">Sim</option>
                                    <option value="0" selected>Não</option>
                                </select>
                            </div>
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
                </div>
            </div>
        </div>
        <div class="nav-footer">
            <button type="button" class="btn btn-sm btn-danger" id="btnRejeitar">
                <i class="bi bi-x-lg me-1"></i>Rejeitar
            </button>
            <button type="button" class="btn btn-sm btn-success" id="btnAprovar">
                <i class="bi bi-check-lg me-1"></i>Aprovar
            </button>
        </div>
    </div>

    <!-- Modal de Avaliação -->
    <div class="modal fade" id="modalAvaliacao" aria-labelledby="modalAvaliacaoTitulo" aria-hidden="true">
        <div class="modal-dialog modal-lg">
            <div class="modal-content">
                <div class="modal-header">
                    <h5 class="modal-title" id="modalAvaliacaoTitulo"></h5>
                    <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Fechar"></button>
                </div>
                <div class="modal-body">
                    <p class="text-muted mb-3" id="modalAvaliacaoDescricao"></p>
                    <label for="avaliacaoObservacao" class="form-label">
                        Observações <span class="text-muted fw-normal" id="avaliacaoObservacaoObrigatoriedade"></span>
                    </label>
                    <textarea class="form-control form-control-sm" id="avaliacaoObservacao" rows="5"></textarea>
                </div>
                <div class="modal-footer">
                    <button type="button" class="btn btn-sm btn-outline-secondary" data-bs-dismiss="modal">
                        Cancelar
                    </button>
                    <button type="button" class="btn btn-sm" id="btnConfirmarAvaliacao"></button>
                </div>
            </div>
        </div>
    </div>

    <!-- Modal de Resultados do Processamento de calculo -->
    <div class="modal fade" id="modalResultadosCalculo" data-backdrop="static" data-keyboard="false" tabindex="-1" role="dialog" aria-labelledby="modalResultadosCalculoLabel" aria-hidden="true">
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