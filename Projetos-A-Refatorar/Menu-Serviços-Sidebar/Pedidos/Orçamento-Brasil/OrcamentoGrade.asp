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
    <title>Lista de Orçamentos</title>
    <meta charset="utf-8"/>
    <meta name="viewport" content="width=device-width, initial-scale=1"/>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.0.2/dist/css/bootstrap.min.css" rel="stylesheet" integrity="sha384-EVSTQN3/azprG1Anm3QDgpJLIm9Nao0Yz1ztcQTwFspd3yD65VohhpuuCOmLASjC" crossorigin="anonymous" />
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.1/font/bootstrap-icons.css" />
    <link href="https://cdn.jsdelivr.net/npm/select2@4.1.0-rc.0/dist/css/select2.min.css" rel="stylesheet" />
    <link href="https://cdn.jsdelivr.net/npm/select2-bootstrap-5-theme@1.3.0/dist/select2-bootstrap-5-theme.min.css" rel="stylesheet" />
    
    <style>
        body {
            background-color: #f5f5f5;
            font-size: .8rem;
            padding-bottom: 20px;
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

        .container {
            max-width: 90%;
        }

        .form-label {
            font-weight: 600;
        }

        .table-wrapper {
            overflow: auto;
            max-height: 70vh;
        }

        #gradeOrcamento th,
        #gradeOrcamento td {
            white-space: nowrap;
            vertical-align: top;
        }

        #gradeOrcamento th:last-child,
        #gradeOrcamento td:last-child {
            position: sticky;
            right: 0;
            z-index: 5;
            white-space: nowrap;
        }
        
        #gradeOrcamento td:last-child {
            background: linear-gradient(to right, rgba(0,0,0,0.1) 0%, transparent 1px),#fff;
            background-origin: border-box;
            background-clip: padding-box;
        }
        
        #gradeOrcamento thead th:last-child {
            background: linear-gradient(to right, rgba(0,0,0,0.1) 0%, transparent 1px), #f5f5f5;
            background-origin: border-box;
            background-clip: padding-box;
            z-index: 10;
        }

        thead {
            vertical-align: middle !important;
            position: sticky; 
            inset-block-start: 0; 
            background: #f5f5f5;
            z-index: 999;
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
    

        .loading-overlay {
            position: absolute;
            top: 0;
            left: 0;
            right: 0;
            bottom: 0;
            background: rgba(255,255,255,0.8);
            display: flex;
            align-items: center;
            justify-content: center;
            z-index: 1000;
        }

        .z-index-9999 {
            z-index: 9999 !important;
        }

        .dropstart .dropdown-toggle::before {
            content: none !important;
        }

        .badge {
            font-size: .7rem;
        }
        .badge-red {
            background-color: #fee2e2;
            color: #991b1b;
        }
        .badge-green {
            background-color: #d1fae5;
            color: #065f46;
        }        
        .badge-purple {
            background-color: #e0e7ff;
            color: #3730a3;
        }
        .badge-pink {
            background-color: #fce7f3;
            color: #9f1239;
        }
        .badge-blue {
            background-color: #cffafe;
            color: #155e75;
        }
        .badge-yellow {
            background-color: #fef3c7;
            color: #92400e;
        }
        .badge-gray {
            background-color: #e2e2e4;
            color: #374151;
        }
         #nivelUsuario {
            font-weight: bold;
        }
    </style>

    <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.0.2/dist/js/bootstrap.bundle.min.js" integrity="sha384-MrcW6ZMFYlzcLA8Nl+NtUVF0sA7MsXsP1UyJoMp4YLEuNSfAP+JcXn/tWtIaxVXM" crossorigin="anonymous"></script>
    <script src="https://cdn.jsdelivr.net/npm/select2@4.1.0-rc.0/dist/js/select2.min.js"></script>
    <script src="/corporativo/components/toast.js?v=<%=Now()%>"></script>
    <script src="/corporativo/components/confirmation.js?v=<%=Now()%>"></script>
    <script src="/corporativo/funcoes/funcoesJAVASCRIPT.js?v=<%=Now()%>"></script>
    <script src="/corporativo/components/dropdownMenu.js?v=<%=Now()%>"></script>

    <script>
        const eStatus = Object.freeze({
            APROVADO_TOTAL: 'Aprovado Total',
            CANCELADO: 'Cancelado',
            EM_ANALISE: 'Em Analise',
            EM_APROVACAO: 'Em Aprovacao',
            EM_DIGITACAO: 'Em Digitacao',
            REPROVADO: 'Reprovado'
        });

        let nivelUsuario = '';
        let dadosOrcamento = [];
        let carregando = false;

        function exibirLoading(exibir) {
            carregando = exibir;

            if (exibir) {
                $('#loadingOverlay').show();
            } else {
                $('#loadingOverlay').hide();
            }
        }

        function obterNivelUsuario() {
            return new Promise((resolve, reject) => {
                $.ajax({
                    url: 'OrcamentoRepositorio.asp?v=<%=Now()%>',
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

        function carregarOrcamentos() { 
            exibirLoading(true);

            const params = {
                acao: 'listar',
                filtroFilial: $('#filtroFilial').val(),
                filtroDataInclusaoDe: $('#filtroDataInclusaoDe').val(),
                filtroDataInclusaoAte: $('#filtroDataInclusaoAte').val(),
                filtroStatus: $('#filtroStatus').val(),
                filtroSituacao: $('#filtroSituacao').val(),
                filtroCodCliente: $('#filtroCodCliente').val(),
                filtroCodLoja: $('#filtroCodLoja').val(),
                filtroNumeroOrcamento: $('#filtroNumeroOrcamento').val(),
            };

            $.ajax({
                url: 'OrcamentoRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: params,
                dataType: 'json',
                success: function(response) {
                    exibirLoading(false);

                    if (!response.sucesso) {
                        Toast.error(response.mensagem);
                        return;
                    }
                
                    dadosOrcamento = response.dados || [];
                    renderizarGrade();
                },
                error: function(xhr, status, error) {
                    exibirLoading(false);
                    console.error('Erro ao pesquisar:', error);
                    Toast.error('Erro ao realizar a pesquisa');
                }
            });
        }

        function renderizarGrade() {
            const $tbody = $('#gradeOrcamento tbody');
            
            $tbody.empty();
            
            if (dadosOrcamento.length === 0) {
                $tbody.html(`
                    <tr>
                        <td colspan="28" class="text-center text-muted py-4">
                            Nenhum registro encontrado
                        </td>
                    </tr>
                `);
                return;
            }

            dadosOrcamento.forEach(function(item) {    
                const statusBadge = obterBadgeStatus(item.status);

                const acoesMap = {
                    'VENDEDOR': obterAcoesVendedor(item),
                    'ANALISTA': obterAcoesAnalista(item),
                    'APROVADOR': obterAcoesAprovador(item),
                }

                let htmlAcoes = acoesMap[nivelUsuario] || '';
                
                $tbody.append(`
                    <tr>
                        <td>${item.numeroOrcamento || '-'}</td>
                        <td>${obterNomePlanta(item.codPlanta)}</td>
                        <td>${item.dataInclusao || '-'}</td>
                        <td>${item.dataValidadeDe || '-'}</td>
                        <td>${item.dataValidadeAte || '-'}</td>
                        <td>${statusBadge}</td>
                        <td>${item.quantidadeItens || 0}</td>
                        <td>${item.codCliente} - ${item.nomeCliente}</td>
                        <td>${formatarCNPJ(item.CNPJ)}</td>
                        <td>${item.cidade || '-'}</td>
                        <td>
                            <div class="truncate-text" title="${item.nomeVendedor || ''}">
                                ${item.nomeVendedor || '-'}
                            </div>
                        </td>
                        ${htmlAcoes}
                    </tr>
                `);
            });
        }

        function novoOrcamento() {
            window.location.href = `OrcamentoFormulario.asp?acao=CRIAR`;
        }

        function visualizarOrcamento(numeroOrcamento) {
            window.location.href = `OrcamentoFormulario.asp?acao=VISUALIZAR&numeroOrcamento=${numeroOrcamento}`;
        }

        function editarOrcamento(numeroOrcamento) {
            window.location.href = `OrcamentoFormulario.asp?acao=EDITAR&numeroOrcamento=${numeroOrcamento}`;
        }

        function analisarOrcamento(numeroOrcamento) {
            window.location.href = `OrcamentoFormulario.asp?acao=ANALISAR&numeroOrcamento=${numeroOrcamento}`;
        }

        function avaliarOrcamento(numeroOrcamento) {
            window.location.href = `OrcamentoFormularioAvaliacao.asp?numeroOrcamento=${numeroOrcamento}`;
        }

        function confirmacaoEnviarParaAnalise(numeroOrcamento) {
            ConfirmationModal.warning({
                message: 'Deseja realmente enviar este orçamento para análise? Essa ação não poderá ser revertida',
                confirmText: 'Confirmar',
                cancelText: 'Cancelar',
                onConfirm: () => {
                    enviarParaAnalise(numeroOrcamento);
                },
                onCancel: () => {}
            });
        }
        
        function enviarParaAnalise(numeroOrcamento) {            
            exibirLoading(true);

            const params = {
                acao: 'enviarParaAnalise',
                numeroOrcamento: numeroOrcamento
            };

            $.ajax({
                url: 'OrcamentoRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: params,
                dataType: 'json',
                success: function(response) {
                    exibirLoading(false);

                    if (!response.sucesso) {
                        Toast.error(response.mensagem || 'Erro ao enviar para análise');
                        return;
                    }

                    Toast.success(response.mensagem || 'Orçamento enviado para análise com sucesso'); 
                    carregarOrcamentos();
                },
                error: function(xhr, status, error) {
                    exibirLoading(false);
                    console.error('Erro ao enviar para análise:', error);
                    Toast.error('Erro ao enviar para análise');
                }
            });
        }

        function confirmacaoEnviarParaAprovacao(numeroOrcamento) {
            ConfirmationModal.warning({
                message: 'Deseja realmente enviar este orçamento para aprovação? Essa ação não poderá ser revertida',
                confirmText: 'Confirmar',
                cancelText: 'Cancelar',
                onConfirm: () => {
                    enviarParaAprovacao(numeroOrcamento);
                },
                onCancel: () => {}
            });
        }
        
        function enviarParaAprovacao(numeroOrcamento) {            
            exibirLoading(true);

            const params = {
                acao: 'enviarParaAprovacao',
                numeroOrcamento: numeroOrcamento
            };

            $.ajax({
                url: 'OrcamentoRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: params,
                dataType: 'json',
                success: function(response) {
                    exibirLoading(false);

                    if (!response.sucesso) {
                        Toast.error(response.mensagem || 'Erro ao enviar para aprovação');
                        return;
                    }

                    Toast.success(response.mensagem || 'Orçamento enviado para aprovação com sucesso'); 
                    carregarOrcamentos();
                },
                error: function(xhr, status, error) {
                    exibirLoading(false);
                    console.error('Erro ao enviar para aprovação:', error);
                    Toast.error('Erro ao enviar para aprovação');
                }
            });
        }

        function confirmacaoCancelamento(numeroOrcamento) {
            ConfirmationModal.danger({
                message: 'Deseja realmente cancelar este orçamento? Essa ação não poderá ser revertida',
                confirmText: 'Confirmar',
                cancelText: 'Cancelar',
                onConfirm: () => {
                    cancelarOrcamento(numeroOrcamento);
                },
                onCancel: () => {}
            });
        }
        
        function cancelarOrcamento(numeroOrcamento) {         
            exibirLoading(true);

            const params = {
                acao: 'cancelar',
                numeroOrcamento: numeroOrcamento
            };

            $.ajax({
                url: 'OrcamentoRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: params,
                dataType: 'json',
                success: function(response) {
                    exibirLoading(false);

                    if (!response.sucesso) {
                        Toast.error(response.mensagem || 'Erro ao cancelar');
                        return;
                    }

                    Toast.success(response.mensagem || 'Orçamento cancelado com sucesso'); 
                    carregarOrcamentos();
                },
                error: function(xhr, status, error) {
                    exibirLoading(false);
                    console.error('Erro ao cancelar:', error);
                    Toast.error('Erro ao cancelar');
                }
            });
        }

        function confirmacaoDuplicacao(numeroOrcamento, codPlanta) {
            $('#duplicacaoNumeroOrcamento').val(numeroOrcamento);
            $('#duplicacaoCodPlanta').val(codPlanta);
            $('#duplicacaoAplicarReajuste').prop('checked', false);
            $('#duplicacaoOpcoesReajuste').addClass('d-none');
            $('#duplicacaoTipoReajuste').val('');
            $('#duplicacaoPercentualReajuste').val('');

            const modalDuplicacao = new bootstrap.Modal(document.getElementById('modalConfirmacaoDuplicacao'));
            modalDuplicacao.show();
        }

        function alternarOpcoesReajusteDuplicacao() {
            if ($('#duplicacaoAplicarReajuste').is(':checked')) {
                $('#duplicacaoOpcoesReajuste').removeClass('d-none');
            } else {
                $('#duplicacaoOpcoesReajuste').addClass('d-none');
                $('#duplicacaoTipoReajuste').val('');
                $('#duplicacaoPercentualReajuste').val('');
            }
        }

        function confirmarDuplicacao() {
            const numeroOrcamento = $('#duplicacaoNumeroOrcamento').val();
            const codPlanta = $('#duplicacaoCodPlanta').val();
            const aplicarReajuste = $('#duplicacaoAplicarReajuste').is(':checked');
            const parametrosReajuste = {};

            if (aplicarReajuste) {
                const tipoReajuste = $('#duplicacaoTipoReajuste').val();
                const percentualReajuste = $('#duplicacaoPercentualReajuste').val();
                const percentualReajusteNumerico = parseFloat(String(percentualReajuste).replace(',', '.'));

                if (!tipoReajuste) {
                    Toast.error('Selecione se deseja aumentar ou reduzir os valores');
                    return;
                }

                if (!percentualReajuste || isNaN(percentualReajusteNumerico) || percentualReajusteNumerico < 1 || percentualReajusteNumerico > 100) {
                    Toast.error('Informe um percentual de reajuste entre 1 e 100');
                    return;
                }

                parametrosReajuste.tipoReajuste = tipoReajuste;
                parametrosReajuste.percentualReajuste = percentualReajusteNumerico;
            }

            const modalDuplicacao = bootstrap.Modal.getInstance(document.getElementById('modalConfirmacaoDuplicacao'));
            modalDuplicacao.hide();

            duplicarOrcamento(numeroOrcamento, codPlanta, parametrosReajuste);
        }
        
        function duplicarOrcamento(numeroOrcamento, codPlanta, parametrosReajuste) {         
            exibirLoading(true);

            const params = {
                acao: 'duplicar',
                codPlanta: codPlanta,
                numeroOrcamento: numeroOrcamento,
                ...(parametrosReajuste || {})
            };

            $.ajax({
                url: 'OrcamentoRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: params,
                dataType: 'json',
                success: function(response) {
                    exibirLoading(false);

                    if (!response.sucesso) {
                        Toast.error(response.mensagem || 'Erro ao duplicar');
                        return;
                    }

                    Toast.success(response.mensagem || 'Orçamento cancelado com sucesso'); 
                    carregarOrcamentos();
                },
                error: function(xhr, status, error) {
                    exibirLoading(false);
                    console.error('Erro ao duplicar:', error);
                    Toast.error('Erro ao duplicar');
                }
            });
        }

        function confirmacaoDevolucao(numeroOrcamento) {
            ConfirmationModal.warning({
                message: 'Deseja realmente devolver este orçamento? Essa ação não poderá ser revertida',
                confirmText: 'Confirmar',
                cancelText: 'Cancelar',
                onConfirm: () => {
                    devolver(numeroOrcamento);
                },
                onCancel: () => {}
            });
        }
        
        function devolver(numeroOrcamento) {            
            exibirLoading(true);

            const params = {
                acao: 'devolver',
                numeroOrcamento: numeroOrcamento
            };

            $.ajax({
                url: 'OrcamentoRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: params,
                dataType: 'json',
                success: function(response) {
                    exibirLoading(false);

                    if (!response.sucesso) {
                        Toast.error(response.mensagem || 'Erro devolver');
                        return;
                    }

                    Toast.success(response.mensagem || 'Orçamento devolvido com sucesso'); 
                    carregarOrcamentos();
                },
                error: function(xhr, status, error) {
                    exibirLoading(false);
                    console.error('Erro devolver:', error);
                    Toast.error('Erro devolver');
                }
            });
        }

        function pesquisarCliente() {
            const codPlanta = $('#filtroFilial').val();
            if(!codPlanta) {
                Toast.warning("Para pesquisar por cliente é necessário selecionar uma filial")
                return;
            }

            const baseUrl = `${location.protocol}//${location.host}`;
            const url = `${baseUrl}/corporativo/buscaCLIENTE.asp?id1=filtroCodCliente&id2=filtroCodLoja&id3=nomeCliente&id6=${codPlanta}`;
            window.open(url, '', "width=650,height=400");
        }

        function limparFiltros() {
            $('#filtroFilial').val('').trigger('change');
            $('#filtroDataInclusaoDe').val('');
            $('#filtroDataInclusaoAte').val('');
            $('#filtroStatus').val('').trigger('change');
            $('#filtroSituacao').val('').trigger('change');
            $('#filtroCodCliente').val('');
            $('#filtroCodLoja').val('');
            $('#filtroNumeroOrcamento').val('');

            dadosOrcamento = [];
            renderizarGrade();
        }

        function obterBadgeStatus(status) {
            switch(status) {
                case eStatus.EM_ANALISE:
                    return '<span class="badge badge-yellow">EM ANÁLISE</span>';
                case eStatus.EM_APROVACAO:
                    return '<span class="badge badge-purple">EM APROVAÇÃO</span>';
                case eStatus.EM_DIGITACAO:
                    return '<span class="badge badge-blue">EM DIGITAÇÃO</span>';
                case eStatus.APROVADO_TOTAL:
                    return '<span class="badge badge-green">APROVADO TOTAL</span>';
                case eStatus.CANCELADO:
                    return '<span class="badge badge-red">CANCELADO</span>';
                case eStatus.REPROVADO:
                    return '<span class="badge badge-red">REPROVADO</span>';
                default:
                    return '-';
            }
        }

        function imprimirCartaProposta(numeroOrcamento) {
            $.ajax({
                url: 'OrcamentoRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: {
                    acao: "obterCartaProposta",
                    numeroOrcamento: numeroOrcamento
                },
                dataType: 'json',
                success: function(response) {
                    if (!response.sucesso) {
                        Toast.error(response.mensagem || 'Erro ao buscar carta.');
                    }

                    const janela = window.open('', '_blank');
                    janela.document.write(response.dados);
                    janela.document.close();

                    janela.onload = () => {
                        janela.focus();
                        janela.print();
                        janela.close();
                    };
                },
                error: function(xhr, status, error) {
                    Toast.error('Erro ao buscar carta.');
                }
            });
        }

        function enviarCartaProposta(numeroOrcamento) {
            $.ajax({
                url: 'OrcamentoRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: {
                    acao: "enviarCartaProposta",
                    numeroOrcamento: numeroOrcamento
                },
                dataType: 'json',
                success: function(response) {
                    if (!response.sucesso) {
                        Toast.error(response.mensagem || 'Erro ao buscar carta.');
                    }

                    Toast.success(response.mensagem || 'Carta enviada com sucesso.');
                },
                error: function(xhr, status, error) {
                    Toast.error('Erro ao buscar carta.');
                }
            });
        }

        function obterAcoesVendedor(item) {
            const itens = [
                `<a class="dropdown-item" onclick="visualizarOrcamento('${item.numeroOrcamento}')"><i class="bi bi-eye me-1"></i> Visualizar</a>`,
            ];

            if(item.status == eStatus.EM_DIGITACAO) {
                itens.push(
                    `<a class="dropdown-item" onclick="editarOrcamento('${item.numeroOrcamento}')"><i class="bi bi-pencil-square me-1"></i> Editar</a>`,
                    `<a class="dropdown-item" onclick="confirmacaoCancelamento('${item.numeroOrcamento}')"><i class="bi bi-ban me-1"></i> Cancelar</a>`,
                )
            }

            if (item.status == eStatus.APROVADO_TOTAL) {
                itens.push(`<a class="dropdown-item" onclick="imprimirCartaProposta('${item.numeroOrcamento}')"><i class="bi bi-printer me-1"></i> Imprimir Carta Proposta</a>`)
            }

            if (item.status == eStatus.EM_DIGITACAO) {
                itens.push(`<a class="dropdown-item" onclick="confirmacaoEnviarParaAnalise('${item.numeroOrcamento}')"><i class="bi bi-clipboard-check me-1"></i> Enviar para Análise</a>`)
            }

            itens.push(
                `<a class="dropdown-item" onclick="confirmacaoDuplicacao('${item.numeroOrcamento}', '${item.codPlanta}')"><i class="bi bi-copy me-1"></i> Duplicar</a>`,
                `<hr class="dropdown-divider">`,
                `<a class="dropdown-item" onclick="abrirModalHistoricoAprovacoes('${item.numeroOrcamento}')"><i class="bi bi-clock-history me-1"></i> Histórico de Aprovações</a>`,
                `<a class="dropdown-item" onclick="abrirModalHistoricoMensagens('${item.numeroOrcamento}')"><i class="bi bi-clock-history me-1"></i> Histórico de Mensagens</a>`
            )

            return `
                <td style="text-align: center">
                    <dropdown-menu label="Ações">
                        ${itens.join('')}
                    </dropdow-menu>
                </td>
            `
        }

        function obterAcoesAnalista(item) {
            const itens = [
                `<a class="dropdown-item" onclick="visualizarOrcamento('${item.numeroOrcamento}')"><i class="bi bi-eye me-1"></i> Visualizar</a>`,
            ];
                
            if (item.status == eStatus.EM_ANALISE || item.status == eStatus.REPROVADO) {
                itens.push(
                    `<a class="dropdown-item" onclick="analisarOrcamento('${item.numeroOrcamento}')"><i class="bi bi-search me-1"></i> Analisar</a>`,
                    `<a class="dropdown-item" onclick="confirmacaoDevolucao('${item.numeroOrcamento}')"><i class="bi bi-arrow-counterclockwise me-1"></i> Devolver</a>`,
                    `<a class="dropdown-item" onclick="confirmacaoEnviarParaAprovacao('${item.numeroOrcamento}')"><i class="bi bi-clipboard-check me-1"></i> Enviar para Aprovação</a>`,
                );
            }

            if (item.status == eStatus.APROVADO_TOTAL) {
                itens.push(  
                    `<a class="dropdown-item" onclick="imprimirCartaProposta('${item.numeroOrcamento}')"><i class="bi bi-printer me-1"></i> Imprimir Carta Proposta</a>`,
                    `<a class="dropdown-item" onclick="enviarCartaProposta('${item.numeroOrcamento}')"><i class="bi bi-box-arrow-up-right me-1"></i> Enviar Carta Proposta</a>`
                )
            }

            itens.push(
                `<hr class="dropdown-divider">`,
                `<a class="dropdown-item" onclick="abrirModalHistoricoAprovacoes('${item.numeroOrcamento}')"><i class="bi bi-clock-history me-1"></i> Histórico de Aprovações</a>`,
                `<a class="dropdown-item" onclick="abrirModalHistoricoMensagens('${item.numeroOrcamento}')"><i class="bi bi-clock-history me-1"></i> Histórico de Mensagens</a>`,
            )

            return `
                <td style="text-align: center">
                    <dropdown-menu label="Ações">
                        ${itens.join('')}
                    </dropdow-menu>
                </td>
            `
        }

        function obterAcoesAprovador(item) {
            const itens = [
                `<a class="dropdown-item" onclick="visualizarOrcamento('${item.numeroOrcamento}')"><i class="bi bi-eye me-1"></i> Visualizar</a>`,
                `<a class="dropdown-item" onclick="avaliarOrcamento('${item.numeroOrcamento}')"><i class="bi bi-clipboard-data me-1"></i> Aprovar ou Rejeitar</a>`,
                `<hr class="dropdown-divider">`,
                `<a class="dropdown-item" onclick="abrirModalHistoricoAprovacoes('${item.numeroOrcamento}')"><i class="bi bi-clock-history me-1"></i> Histórico de Aprovações</a>`,
                `<a class="dropdown-item" onclick="abrirModalHistoricoMensagens('${item.numeroOrcamento}')"><i class="bi bi-clock-history me-1"></i> Histórico de Mensagens</a>`,
           ];

            return `
                <td style="text-align: center">
                    <dropdown-menu label="Ações">
                        ${itens.join('')}
                    </dropdow-menu>
                </td>
            `
        }

        function abrirModalHistoricoAprovacoes(numeroOrcamento) {
            $('#numeroOrcamentoHistoricoAprovacao').text(numeroOrcamento);
            const modal = new bootstrap.Modal(document.getElementById('modalHistoricoAprovacoes'));
            modal.show();

            const $tbody = $('#gradeHistoricoAprovacao tbody');
            $tbody.empty();
     
            $.ajax({
                url: 'OrcamentoRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: {
                    acao: "listarHistoricoAprovacoes",
                    numeroOrcamento: numeroOrcamento
                },
                dataType: 'json',
                success: function(response) {
                    if (!response.sucesso) {
                        Toast.error(response.mensagem || 'Erro ao buscar aprovações.');

                        $tbody.html(`
                            <tr>
                                <td colspan="28" class="text-center text-muted py-4">
                                    Nenhum registro encontrado
                                </td>
                            </tr>
                        `);
                        return;
                    }

                    if (response.dados.length === 0) {
                        $tbody.html(`
                            <tr>
                                <td colspan="28" class="text-center text-muted py-4">
                                    Nenhum registro encontrado
                                </td>
                            </tr>
                        `);
                        return;
                    }

                    response.dados.forEach(function(item) {    
                        $tbody.append(`
                            <tr>
                                <td>${item.nomeAprovador || '-'}</td>
                                <td>${item.nivel || '-'}</td>
                                ${criarCelulaDetalhe(item.justificativa, 'Descrição da resposta')}
                                <td>${item.dataAprovacao || '-'}</td>
                            </tr>
                        `);
                    });
                },
                error: function(xhr, status, error) {
                    console.error('Erro ao buscar histórico de aprovações:', error);
                    $tbody.html(`
                        <tr>
                            <td colspan="28" class="text-center text-muted py-4">
                                Nenhum registro encontrado
                            </td>
                        </tr>
                    `);
                }
            });
        }
   
        function abrirModalHistoricoMensagens(numeroOrcamento) {
            $('#numeroOrcamentoHistoricoMensagem').text(numeroOrcamento);
            const modal = new bootstrap.Modal(document.getElementById('modalHistoricoMensagens'));
            modal.show();

            const $tbody = $('#gradeHistoricoMensagem tbody');
            $tbody.empty();
     
            $.ajax({
                url: 'OrcamentoRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: {
                    acao: "listarHistoricoMensagens",
                    numeroOrcamento: numeroOrcamento
                },
                dataType: 'json',
                success: function(response) {
                    if (!response.sucesso) {
                        Toast.error(response.mensagem || 'Erro ao buscar aprovações.');

                        $tbody.html(`
                            <tr>
                                <td colspan="28" class="text-center text-muted py-4">
                                    Nenhum registro encontrado
                                </td>
                            </tr>
                        `);
                        return;
                    }

                    if (response.dados.length === 0) {
                        $tbody.html(`
                            <tr>
                                <td colspan="28" class="text-center text-muted py-4">
                                    Nenhum registro encontrado
                                </td>
                            </tr>
                        `);
                        return;
                    }

                    response.dados.forEach(function(item) {    
                        $tbody.append(`
                            <tr>
                                <td>${item.nomeUsuarioInclusao || '-'}</td>
                                ${criarCelulaDetalhe(item.descricao, 'Mensagem')}
                                <td>${item.dataInclusao || '-'}</td>
                            </tr>
                        `);
                    });
                },
                error: function(xhr, status, error) {
                    console.error('Erro ao buscar histórico de aprovações:', error);
                    $tbody.html(`
                        <tr>
                            <td colspan="28" class="text-center text-muted py-4">
                                Nenhum registro encontrado
                            </td>
                        </tr>
                    `);
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

        function inicializarSelect2() {
            $('.select2').select2({
                theme: 'bootstrap-5',
                width: '100%',
                allowClear: true,
                dropdownCssClass: "z-index-9999",
                placeholder: function() {
                    return $(this).data('placeholder');
                }
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

            $('#nivelUsuario').text(nivelUsuario);

            if(nivelUsuario == 'VENDEDOR') {
                $('#btnNovoOrcamento').removeClass('d-none');
            }

            if(nivelUsuario == 'ANALISTA') {
                $('#btnDesativarOrcamentos').removeClass('d-none');
            }

            inicializarSelect2();

            const urlParams = new URLSearchParams(window.location.search);

            if (urlParams.has('numeroOrcamento')) 
                $('#filtroNumeroOrcamento').val(urlParams.get('numeroOrcamento'));

            if (urlParams.has('codPlanta')) 
                $('#filtroFilial').val(urlParams.get('codPlanta')).trigger('change');

            carregarOrcamentos();
        });
    </script>
</head>
<body>
    <div class="container my-4">
        <div class="my-4">
            <h3>Lista de Orçamentos</h3>
        </div>

        <div class="card p-3 mb-4">
            <div class="row g-3">
                <div class="col-12 col-md-4 col-lg-3">
                    <label for="filtroFilial" class="form-label">Filial</label>
                    <select required class="form-select form-select-sm select2" id="filtroFilial" name="filtroFilial" data-placeholder="Todas as filiais">
                        <option value="">Todas as filiais</option>
                        <option value="00">Pirapetinga</option>
                        <option value="04">Uberaba</option>
                        <option value="05">Bento</option>
                        <option value="06">Maranguape</option>
                        <option value="07">Saquarema</option>
                    </select>
                </div>

                <div class="col-12 col-md-4 col-lg-3">
                    <label for="filtroDataInclusaoDe" class="form-label">Data Inclusão De</label>
                    <input type="date" class="form-control form-control-sm" id="filtroDataInclusaoDe" name="filtroDataInclusaoDe">
                </div>

                <div class="col-12 col-md-4 col-lg-3">
                    <label for="filtroDataInclusaoAte" class="form-label">Data Inclusão Até</label>
                    <input type="date" class="form-control form-control-sm" id="filtroDataInclusaoAte" name="filtroDataInclusaoAte">
                </div>
                
                <div class="col-12 col-md-4 col-lg-3">
                    <label for="filtroCodCliente" class="form-label">Cód. Cliente</label>
                    <div class="input-group input-group-sm">
                        <button type="button" class="btn btn-sm btn-secondary" onclick="pesquisarCliente()">
                            Buscar <i class="bi bi-box-arrow-up-right"></i>
                        </button>
                        <input type="text" class="form-control form-control-sm" id="filtroCodCliente" name="filtroCodCliente" placeholder="Código do cliente">
                    </div>
                    <small class="text-muted">Clique no botão para buscar ou digite o código</small>
                </div>

                <div class="col-12 col-md-4 col-lg-3">
                    <label for="filtroCodLoja" class="form-label">Cód. Loja</label>
                    <input type="text" class="form-control form-control-sm" id="filtroCodLoja" name="filtroCodLoja" placeholder="Código da loja">
                </div>
                
                <div class="col-12 col-md-4 col-lg-3">
                    <label for="filtroStatus" class="form-label">Status</label>
                    <select required class="form-select form-select-sm select2" id="filtroStatus" name="filtroStatus" data-placeholder="Todos">
                        <option value="">Todos</option>
                        <option value="Em Digitacao">Em Digitação</option>
                        <option value="Em Analise">Em Análise</option>
                        <option value="Em Aprovacao">Em Aprovação</option>
                        <option value="Aprovado Total">Aprovado Total</option>
                        <option value="Reprovado">Reprovado</option>
                    </select>
                </div>

                <div class="col-12 col-md-4 col-lg-3">
                    <label for="filtroNumeroOrcamento" class="form-label">Nº do Orçamento</label>
                    <input type="number" class="form-control form-control-sm" id="filtroNumeroOrcamento" name="filtroNumeroOrcamento" placeholder="Nº do Orçamento">
                </div>

                <div class="col-12 d-flex justify-content-end align-items-end gap-2 mt-4">
                    <button onclick="limparFiltros()" type="button" class="btn btn-sm btn-outline-secondary" id="btnLimpar">
                        <i class="bi bi-x-lg me-1"></i> Limpar
                    </button>
                    <button onclick="carregarOrcamentos()" type="button" class="btn btn-sm btn-primary" id="btnPesquisar">
                        <i class="bi bi-search me-1"></i> Pesquisar
                    </button>
                </div>
            </div>
        </div>

        <div class="alert alert-primary alert-dismissible py-2 px-3 mb-4 d-flex gap-1 align-items-center" role="alert" style="font-size: 0.85rem;">
            <i class="bi bi-info-circle me-2"></i>
            Visualizando como: <span id="nivelUsuario"></span>
        </div>

        <div class="card ">
            <div class="gap-2 p-3 border-bottom">
                <button id="btnNovoOrcamento" class="d-none btn btn-sm btn-primary" onclick="novoOrcamento()">
                    <i class="bi bi-plus-lg"></i>
                    Novo
                </button>
                <button id="btnDesativarOrcamentos" type="button" class="d-none btn  btn-primary btn-sm" onclick="window.location.href = 'OrcamentoDesativar.asp'">
                    <i class="bi bi-box-arrow-up-right"></i>
                    Desativar Orçamentos
                </button>
            </div>

            <div class="table-wrapper">
                <div id="loadingOverlay" class="loading-overlay" style="display: none;">
                    <div class="spinner-border text-primary" role="status">
                        <span class="visually-hidden">Carregando...</span>
                    </div>
                </div>

                <table class="table table-hover mb-0" id="gradeOrcamento">
                    <thead>
                        <tr>
                            <th>#</th>
                            <th>Planta</th>
                            <th>Data Inclusão</th>
                            <th>Data Validade De</th>
                            <th>Data Validate Até</th>
                            <th>Status</th>
                            <th>Qtd. Itens</th>
                            <th>Cliente</th>
                            <th>CNPJ</th>
                            <th>Cidade</th>
                            <th>Vendedor</th>
                            <th style="width: 120px; text-align: center">Ações</th>
                        </tr>
                    </thead>
                    <tbody>
                        <tr>
                            <td colspan="29" class="text-center text-muted py-4">
                                Nenhum registro encontrado
                            </td>
                        </tr>
                    </tbody>
                </table>
            </div>
        </div>
    </div>
    
    <!-- Modal de Confirmação de Duplicação -->
    <div class="modal fade" id="modalConfirmacaoDuplicacao" tabindex="-1" aria-labelledby="modalConfirmacaoDuplicacaoLabel" aria-hidden="true">
        <div class="modal-dialog">
            <div class="modal-content">
                <div class="modal-header">
                    <h5 class="modal-title" id="modalConfirmacaoDuplicacaoLabel">Duplicar Orçamento</h5>
                    <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Fechar"></button>
                </div>
                <div class="modal-body">
                    <input type="hidden" id="duplicacaoNumeroOrcamento">
                    <input type="hidden" id="duplicacaoCodPlanta">

                    <p class="mb-3">Deseja realmente duplicar este orçamento? Essa ação não poderá ser revertida.</p>

                    <div class="form-check mb-3">
                        <input class="form-check-input" type="checkbox" id="duplicacaoAplicarReajuste" onchange="alternarOpcoesReajusteDuplicacao()">
                        <label class="form-check-label" for="duplicacaoAplicarReajuste">
                            Aplicar reajuste de <strong>Valor Unitário</strong> nos itens antes de duplicar
                        </label>
                    </div>

                    <div id="duplicacaoOpcoesReajuste" class="d-none">
                        <div class="mb-3">
                            <label for="duplicacaoTipoReajuste" class="form-label">Tipo de reajuste</label>
                            <select class="form-select form-select-sm" id="duplicacaoTipoReajuste">
                                <option value="">Selecione</option>
                                <option value="aumentar">Aumentar valores</option>
                                <option value="reduzir">Reduzir valores</option>
                            </select>
                        </div>

                        <div class="mb-0">
                            <label for="duplicacaoPercentualReajuste" class="form-label">Percentual de reajuste</label>
                            <div class="input-group input-group-sm">
                                <input type="number" class="form-control form-control-sm" id="duplicacaoPercentualReajuste" min="1" max="100" step="0.01" placeholder="Informe um valor de 1 a 100">
                                <span class="input-group-text">%</span>
                            </div>
                            <small class="text-muted">Ex.: 10 aumenta ou reduz 10% em cada item.</small>
                        </div>
                    </div>
                </div>
                <div class="modal-footer">
                    <button type="button" class="btn btn-sm btn-secondary" data-bs-dismiss="modal">Cancelar</button>
                    <button type="button" class="btn btn-sm btn-warning" onclick="confirmarDuplicacao()">Confirmar</button>
                </div>
            </div>
        </div>
    </div>

    <!-- Modal de Histórico de aprovações -->
    <div class="modal fade" id="modalHistoricoAprovacoes" tabindex="-1" aria-labelledby="modalHistoricoAprovacoesLabel" aria-hidden="true">
        <div class="modal-dialog modal-xl">
            <div class="modal-content">
                <div class="modal-header">
                    <h5 class="modal-title" id="modalHistoricoAprovacoesTitulo">
                        Histórico de Aprovações #<span id="numeroOrcamentoHistoricoAprovacao"></span>
                    </h5>
                    <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Fechar"></button>
                </div>
                <div class="modal-body">
                    <table class="table table-hover mb-0" id="gradeHistoricoAprovacao">
                        <thead>
                            <tr>
                                <th>Aprovador</th>
                                <th>Nível</th>
                                <th>Justificativa</th>
                                <th>Data Aprovação</th>
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
    
    <!-- Modal de Histórico de mensagems -->
    <div class="modal fade" id="modalHistoricoMensagens" tabindex="-1" aria-labelledby="modalHistoricoMensagensLabel" aria-hidden="true">
        <div class="modal-dialog modal-xl">
            <div class="modal-content">
                <div class="modal-header">
                    <h5 class="modal-title" id="modalHistoricoMensagensTitulo">
                        Histórico de Mensagens #<span id="numeroOrcamentoHistoricoMensagem"></span>
                    </h5>
                    <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Fechar"></button>
                </div>
                <div class="modal-body">
                    <table class="table table-hover mb-0" id="gradeHistoricoMensagem">
                        <thead>
                            <tr>
                                <th>Usuário</th>
                                <th>Mensagem</th>
                                <th>Data Inclusao</th>
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