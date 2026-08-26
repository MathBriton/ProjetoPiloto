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
    <title>Lote Piloto</title>
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

        #gradeLotePiloto th,
        #gradeLotePiloto td {
            white-space: nowrap;
            vertical-align: top;
        }

        #gradeLotePiloto th:last-child,
        #gradeLotePiloto td:last-child {
            position: sticky;
            right: 0;
            z-index: 5;
            white-space: nowrap;
        }
        
        #gradeLotePiloto td:last-child {
            background: linear-gradient(to right, rgba(0,0,0,0.1) 0%, transparent 1px),#fff;
            background-origin: border-box;
            background-clip: padding-box;
        }
        
        #gradeLotePiloto thead th:last-child {
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
        .badge-teal {
            background-color: #ccfbf1;
            color: #134e4a;
        }
        .badge-orange {
            background-color: #ffedd5;
            color: #9a3412;
        }
        .form-switch {
            display: flex;
            align-items: center;
        }
        .form-switch .form-check-input {
            width: 4em;
            height: 1.8em;
            cursor: pointer;
        }
        .form-switch .form-check-label {
            cursor: pointer;
            font-weight: 600;
            color: #495057;
            padding-left: 8px;
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
            AGUARDANDO_PEDIDO_LP: '1',
            CONVERTIDO_VENDA: '2',
            APROVADO_SEM_PEDIDO_CONVERSAO: '3',
            EM_TESTE_CLIENTE: '4',
            REPROVADO_CLIENTE: '5',
            REPROVADO_INTERNAMENTE: '6',
            AGUARDANDO_PRODUCAO: '7',
            DESCONTINUADO: '8',
            SEM_RETORNO_VENDAS: '9',
            PROCESSO_SEGUE: '10',
            VALIDACAO_MAQUINA_NOVA: '11',
            EXCLUIDO: '99'
        });

        const eMotivo = Object.freeze({
            INOVACAO: '1',
            MELHORIA_SIGNIFICATIVA: '2',
            MELHORIA_PALETIZACAO_TRANSPORTE: '3', 
            NOVO_SKU_CLIENTE_ACOMPANHAMENTO_DETALHADO: '4',
            NOVO_SKU_ACOMPANHAMENTO_DETALHADO: '5',
            NOVO_SKU_LAYOUT_IMPRESSAO: '6',
            NOVO_SKU_DIMENSIONAL: '7',
            NOVO_SKU_PADRAO_COR: '8',
            NOVO_SKU_PALETIZACAO_TRANSPORTE: '9',
        });

        let dadosLotePiloto = [];
        let carregando = false;

        function exibirLoading(exibir) {
            carregando = exibir;

            if (exibir) {
                $('#loadingOverlay').show();
            } else {
                $('#loadingOverlay').hide();
            }
        }

        function obterMotivoTraduzido(motivo) {
            const motivosMap = {
                '1': 'INOVAÇÃO',
                '2': 'MELHORIA SIGNIFICATIVA - PROCESSO SW & CLIENTE',
                '3': 'MELHORIA - PALETIZAÇÃO E TRANSPORTE',
                '4': 'NOVO SKU - SW + CLIENTE  - ACOMPANHAMENTO DETALHADO',
                '5': 'NOVO SKU - SW - ACOMPANHAMENTO DETALHADO',
                '6': 'NOVO SKU - LAYOUT DE IMPRESSÃO',
                '7': 'NOVO SKU - DIMENSIONAL',
                '8': 'NOVO SKU - PADRÃO DE COR / CROMIA / RETICULA',
                '9': 'NOVO SKU - PALETIZAÇÃO E TRANSPORTE'
            }

            return motivosMap[motivo] || ''
        }

        function obterTipoLPTraduzido(motivo) {
            const motivosMap = {
                '1': 'NOVO CLIENTE',
                '2': 'INCREMENTO DE VOLUME',
                '3': 'MANUTENCAO DE VOLUME',
            }

            return motivosMap[motivo] || ''
        }

        function carregarStatus() {
            const $select = $('#filtroStatus');
            $select.empty().append('<option value="">Selecione um item</option>');

            return $.ajax({
                url: 'LotePilotoRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: { acao: 'listarStatus' },
                dataType: 'json',
                success: function(response) {
                    if (!response.sucesso) {
                        console.error('Erro ao carregar status: ' + response.mensagem);
                        return;
                    }

                    (response.dados || []).forEach(function(item) {
                        $select.append($('<option>', { value: item.codigo, text: item.descricao }));
                    });

                    $select.trigger('change.select2');
                },
                error: function(xhr, status, error) {
                    console.error('Erro ao carregar status:', error);
                }
            });
        }

        function carregarLotesPiloto() { 
            exibirLoading(true);

            const params = {
                acao: 'listar',
                filtroFilial: $('#filtroFilial').val(),
                filtroDataInclusaoDe: $('#filtroDataInclusaoDe').val(),
                filtroDataInclusaoAte: $('#filtroDataInclusaoAte').val(),
                filtroStatus: $('#filtroStatus').val(),
                filtroSituacao: $('#filtroSituacao').val(),
                filtroCodCliente: $('#filtroCodCliente').val(),
            };

            $.ajax({
                url: 'LotePilotoRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: params,
                dataType: 'json',
                success: function(response) {
                    exibirLoading(false);

                    if (!response.sucesso) {
                        Toast.error(response.mensagem);
                        return;
                    }
                
                    dadosLotePiloto = response.dados || [];
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
            const $tbody = $('#gradeLotePiloto tbody');
            
            $tbody.empty();
            
            if (dadosLotePiloto.length === 0) {
                $tbody.html(`
                    <tr>
                        <td colspan="28" class="text-center text-muted py-4">
                            Nenhum registro encontrado
                        </td>
                    </tr>
                `);
                return;
            }

            dadosLotePiloto.forEach(function(item) {    
                const statusBadge = obterBadgeStatus(item.status);
                
                let htmlAcoes = obterAcoes(item)

                $tbody.append(`
                    <tr>
                        <td>${item.id || '-'}</td>
                        <td>${obterNomePlanta(item.codigoPlanta)}</td>
                        <td>${formatarParaDataBR(item.dataInclusao, false)}</td>
                        <td>${item.dataStatus ? formatarParaDataBR(item.dataStatus, false) : '-'}</td>
                        <td>${statusBadge}</td>
                        <td>${item.codigoCliente} - ${item.nomeCliente}</td>
                        <td>${item.codigoProduto}</td>
                        <td>${item.quantidadePaletes}</td>
                        <td>${item.quantidadePedidos}</td>
                        <td>${obterMotivoTraduzido(item.motivo)}</td>
                        <td>${obterTipoLPTraduzido(item.tipoLP)}</td>
                        <td>
                            <div class="truncate-text" title="${item.nomeUsuarioInclusao || ''}">
                                ${item.nomeUsuarioInclusao || '-'}
                            </div>
                        </td>
                        ${htmlAcoes}
                    </tr>
                `);
            });
        }

        function novoLotePiloto() {
            window.location.href = `LotePilotoFormulario.asp?acao=CRIAR`;
        }

        function visualizarLotePiloto(id) {
            window.location.href = `LotePilotoFormulario.asp?acao=VISUALIZAR&id=${id}`;
        }

        function editarLotePiloto(id) {
            window.location.href = `LotePilotoFormulario.asp?acao=EDITAR&id=${id}`;
        }

        /* tipo: NOVO_LP ou MESMO_PRODUTO */
        function continuarLotePiloto(id, tipo) {
            window.location.href = `LotePilotoFormulario.asp?acao=CONTINUAR&id=${id}&tipo=${tipo}`;
        }

        function abrirModalDetalhesPedido(id, codigoPlanta, codigoCliente, codigoLoja, codigoProduto, quantidadePedidos) {
            const modal = new bootstrap.Modal(document.getElementById('modalDetalhesPedido'));
            const $conteudo = $('#conteudoDetalhesPedido');
            $('#idDetalhesPedido').text(id || '-');

            $('#clienteDetalhesPedido').text(codigoCliente || '-');
            $('#produtoDetalhesPedido').text(codigoProduto || '-');
            $conteudo.html(`
                <div class="text-center text-muted py-4">
                    <div class="spinner-border spinner-border-sm text-primary me-2" role="status">
                        <span class="visually-hidden">Carregando...</span>
                    </div>
                    Carregando detalhes...
                </div>
            `);
            modal.show();

            $.ajax({
                url: 'LotePilotoRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: {
                    acao: 'obterDetalhesPedido',
                    codigoPlanta: codigoPlanta,
                    codigoCliente: codigoCliente,
                    codigoLoja: codigoLoja,
                    codigoProduto: codigoProduto,
                    quantidadePedidos: quantidadePedidos
                },
                dataType: 'json',
                success: function(response) {
                    if (!response.sucesso) {
                        Toast.error(response.mensagem || 'Erro ao buscar detalhes do pedido.');
                        $conteudo.html(`
                            <div class="text-center text-muted py-4">
                                Nenhum registro encontrado
                            </div>
                        `);
                        return;
                    }

                    const item = response.dados || {};
                    $('#clienteDetalhesPedido').text(`${item.codigoCliente || codigoCliente || '-'} - ${item.nomeCliente || '-'}`);
                    $('#produtoDetalhesPedido').text(item.codigoProduto || codigoProduto || '-');

                    $conteudo.html(`
                        <div class="table-responsive">
                            <table class="table table-bordered mb-0">
                                <thead>
                                    <tr>
                                        <th>Tipo</th>
                                        <th class="text-end">Pedidos</th>
                                        <th class="text-end">Quantidade</th>
                                    </tr>
                                </thead>
                                <tbody>
                                    <tr>
                                        <td>Com lote piloto</td>
                                        <td class="text-end">${item.pedidosComLotePiloto || '0'}</td>
                                        <td class="text-end">${item.quantidadeComLotePiloto || '0'}</td>
                                    </tr>
                                    <tr>
                                        <td>Sem lote piloto</td>
                                        <td class="text-end">${item.pedidosSemLotePiloto || '0'}</td>
                                        <td class="text-end">${item.quantidadeSemLotePiloto || '0'}</td>
                                    </tr>
                                </tbody>
                                <tfoot>
                                    <tr class="fw-bold">
                                        <td>Total</td>
                                        <td class="text-end">${item.totalPedidos || '0'}</td>
                                        <td class="text-end">${item.totalQuantidade || '0'}</td>
                                    </tr>
                                </tfoot>
                            </table>
                        </div>
                    `);
                },
                error: function(xhr, status, error) {
                    console.error('Erro ao buscar detalhes do pedido:', error);
                    Toast.error('Erro ao buscar detalhes do pedido.');
                    $conteudo.html(`
                        <div class="text-center text-muted py-4">
                            Nenhum registro encontrado
                        </div>
                    `);
                }
            });
        }

        function abrirModalHistoricoStatus(id) {
            $('#idHistoricoStatus').text(id);
            const modal = new bootstrap.Modal(document.getElementById('modalHistoricoStatus'));
            modal.show();

            const $tbody = $('#gradeHistoricoStatus tbody');
            $tbody.empty();
                 
            $.ajax({
                url: 'LotePilotoRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: {
                    acao: "listarHistoricoStatus",
                    id: id
                },
                dataType: 'json',
                success: function(response) {
                    if (!response.sucesso) {
                        Toast.error(response.mensagem || 'Erro ao buscar histórico.');

                        $tbody.html(`
                            <tr>
                                <td colspan="6" class="text-center text-muted py-4">
                                    Nenhum registro encontrado
                                </td>
                            </tr>
                        `);
                        return;
                    }

                    if (response.dados.length === 0) {
                        $tbody.html(`
                            <tr>
                                <td colspan="6" class="text-center text-muted py-4">
                                    Nenhum registro encontrado
                                </td>
                            </tr>
                        `);
                        return;
                    }


                    response.dados.forEach(function(item) { 
                        $tbody.append(`
                            <tr>
                                <td>${item.id || '-'}</td>
                                <td>${item.statusAnterior || '-'}</td>
                                <td>${item.statusNovo || '-'}</td>
                                <td>${item.alteradoPor || '-'}</td>
                                <td>${item.nomeUsuarioInclusao || '-'}</td>
                                <td>${item.dataInclusao}</td>
                            </tr>
                        `);
                    });
                },
                error: function(xhr, status, error) {
                    console.error('Erro ao buscar histórico de respostas:', error);
                    $tbody.html(`
                        <tr>
                            <td colspan="6" class="text-center text-muted py-4">
                                Nenhum registro encontrado
                            </td>
                        </tr>
                    `);
                }
            });
        }

        function abrirModalHistoricoJustificativas(id) {
            $('#idHistoricoJustificativas').text(id);
            const modal = new bootstrap.Modal(document.getElementById('modalHistoricoJustificativas'));
            modal.show();

            const $tbody = $('#gradeHistoricoJustificativas tbody');
            $tbody.empty();
                 
            $.ajax({
                url: 'LotePilotoRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: {
                    acao: "listarHistoricoJustificativas",
                    id: id
                },
                dataType: 'json',
                success: function(response) {
                    if (!response.sucesso) {
                        Toast.error(response.mensagem || 'Erro ao buscar histórico.');

                        $tbody.html(`
                            <tr>
                                <td colspan="6" class="text-center text-muted py-4">
                                    Nenhum registro encontrado
                                </td>
                            </tr>
                        `);
                        return;
                    }

                    if (response.dados.length === 0) {
                        $tbody.html(`
                            <tr>
                                <td colspan="6" class="text-center text-muted py-4">
                                    Nenhum registro encontrado
                                </td>
                            </tr>
                        `);
                        return;
                    }


                    response.dados.forEach(function(item) { 
                        $tbody.append(`
                            <tr>
                                <td>${item.id || '-'}</td>
                                <td>${item.tipoJustificativa || '-'}</td>
                                <td>${item.descricaoJustificativa || '-'}</td>
                                <td>${item.nomeUsuarioInclusao || '-'}</td>
                                <td>${item.dataInclusao}</td>
                            </tr>
                        `);
                    });
                },
                error: function(xhr, status, error) {
                    console.error('Erro ao buscar histórico de respostas:', error);
                    $tbody.html(`
                        <tr>
                            <td colspan="6" class="text-center text-muted py-4">
                                Nenhum registro encontrado
                            </td>
                        </tr>
                    `);
                }
            });
        }

        function limparFiltros() {
            $('#filtroFilial').val('').trigger('change');
            $('#filtroDataInclusaoDe').val('');
            $('#filtroDataInclusaoAte').val('');
            $('#filtroStatus').val('').trigger('change');
            $('#filtroSituacao').val('').trigger('change');
            $('#filtroCodCliente').val('');

            dadosLotePiloto = [];
            renderizarGrade();
        }

        function obterBadgeStatus(status) {
            switch(status) {
                case eStatus.AGUARDANDO_PEDIDO_LP:
                    return '<span class="badge badge-yellow">AGUARDANDO PEDIDO LP</span>';
                case eStatus.AGUARDANDO_PRODUCAO:
                    return '<span class="badge badge-purple">AGUARDANDO PRODUCAO</span>';
                case eStatus.CONVERTIDO_VENDA:
                    return '<span class="badge badge-purple">CONVERTIDO EM VENDA</span>';
                case eStatus.APROVADO_SEM_PEDIDO_CONVERSAO:
                    return '<span class="badge badge-blue">APROVADO SEM PEDIDO DE CONVERSAO</span>';
                case eStatus.EM_TESTE_CLIENTE:
                    return '<span class="badge badge-green">EM TESTE CLIENTE</span>';
                case eStatus.REPROVADO_CLIENTE:
                    return '<span class="badge badge-pink">REPROVADO PELO CLIENTE</span>';
                case eStatus.REPROVADO_INTERNAMENTE:
                    return '<span class="badge badge-red">REPROVADO INTERNAMENTE/RETRABALHO</span>';
                case eStatus.DESCONTINUADO:
                    return '<span class="badge badge-orange">DESCONTINUADO</span>';
                case eStatus.SEM_RETORNO_VENDAS:
                    return '<span class="badge badge-gray">SEM RETORNO VENDAS/CLIENTE</span>';
                case eStatus.PROCESSO_SEGUE:
                    return '<span class="badge badge-yellow">PROCESSO SEGUE EM NOVO PRODUTO/FICHA TECNICA</span>';
                case eStatus.VALIDACAO_MAQUINA_NOVA:
                    return '<span class="badge badge-green">VALIDACAO MAQUINA NOVA/HOMOLOGACAO</span>';
                case eStatus.EXCLUIDO:
                    return '<span class="badge badge-red">EXCLUÍDO</span>';
                default:
                    return '-';
            }
        }

        function obterAcoes(item) {
            const itens = [
            ];

            if(item.status != eStatus.EXCLUIDO) {
                itens.push(
                    `<a class="dropdown-item" onclick="editarLotePiloto('${item.id}')"><i class="bi bi-pencil-square me-1"></i> Editar</a>`,
                    `<a class="dropdown-item" onclick="abrirModalExclusao('${item.id}', '${item.status}', '${item.codigoPlanta}', '${item.codigoProduto}')"><i class="bi bi-trash me-1"></i> Excluir</a>`,
                )
            }

            if(item.status == eStatus.REPROVADO_CLIENTE || item.status == eStatus.PROCESSO_SEGUE) {
                itens.push(
                    `<a class="dropdown-item" onclick="continuarLotePiloto('${item.id}', 'NOVO_LP')"><i class="bi bi-arrow-repeat"></i> Continuação - Novo LP</a>`,
                    `<a class="dropdown-item" onclick="continuarLotePiloto('${item.id}', 'MESMO_PRODUTO')"><i class="bi bi-arrow-repeat"></i> Continuação - Mesmo Produto</a>`
                )
            }

            itens.push(
                `<hr class="dropdown-divider">`,
                `<a class="dropdown-item" onclick="abrirModalDetalhesPedido('${item.id}', '${item.codigoPlanta}', '${item.codigoCliente}', '${item.codigoLoja}', '${item.codigoProduto}', '${item.quantidadePedidos}')"><i class="bi bi-file-text"></i> Detalhes de Pedido</a>`,
                `<a class="dropdown-item" onclick="abrirModalHistoricoStatus('${item.id}')"><i class="bi bi-clock-history"></i> Histórico de Status</a>`,
                `<a class="dropdown-item" onclick="abrirModalHistoricoJustificativas('${item.id}')"><i class="bi bi-clock-history"></i> Histórico de Justificativas</a>`
            )

            return `
                <td style="text-align: center">
                    <dropdown-menu label="Ações">
                        ${itens.join('')}
                    </dropdow-menu>
                </td>
            `
        }

        function abrirModalExclusao(id, status, codigoPlanta, codigoProduto) {
            $('#modalExclusaoTituloTexto').text(`Confirmar exclusão lote piloto #${id}`)
            $('#idExclusao').val(id)
            $('#statusExclusao').val(status)
            $('#codigoPlantaExclusao').val(codigoPlanta)
            $('#codigoProdutoExclusao').val(codigoProduto)

            $('#descricaoJustificativa').val('').removeClass('is-invalid');
            $('#feedbackJustificativa').text('');

            const modal = new bootstrap.Modal(document.getElementById('modalExclusao'));
            modal.show();
        }

        function excluirLotePiloto() {
            console.log('excluirLotePiloto')

            const modal = bootstrap.Modal.getInstance(document.getElementById('modalExclusao'));
            const descricaoJustificativa = $.trim($('#descricaoJustificativa').val());

            $('#descricaoJustificativa').removeClass('is-invalid');
            $('#feedbackJustificativa').text('');

            if (!descricaoJustificativa) {
                $('#descricaoJustificativa').addClass('is-invalid');
                $('#feedbackJustificativa').text('Informe uma justificativa para continuar.');
                return;
            }

            modal.hide();
            exibirLoading(true);

            const params = {
                acao: 'excluir',
                id: $('#idExclusao').val(),
                status: $('#statusExclusao').val(),
                codigoPlanta: $('#codigoPlantaExclusao').val(),
                codigoProduto: $('#codigoProdutoExclusao').val(),
                descricaoJustificativa: descricaoJustificativa
            };

            $.ajax({
                url: 'LotePilotoRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: params,
                dataType: 'json',
                success: function(response) {
                    exibirLoading(false);

                    if (!response.sucesso) {
                        Toast.error(response.mensagem);
                        return;
                    }

                    Toast.success(response.mensagem); 
                    carregarLotesPiloto();
                },
                error: function(xhr, status, error) {
                    exibirLoading(false);
                    console.error(error);
                    Toast.error(error);
                }
            });
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
            inicializarSelect2();

            await Promise.all([
                carregarStatus(),
                carregarLotesPiloto()
            ]);
        });
    </script>
</head>
<body>
    <div class="container my-4">
        <div class="my-4">
            <h3>Lote Piloto</h3>
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
                    <input type="text" class="form-control form-control-sm" id="filtroCodCliente" name="filtroCodCliente" placeholder="Código do cliente">
                </div>
                
                <div class="col-12 col-md-4 col-lg-3">
                    <label for="filtroStatus" class="form-label">Status</label>
                    <select required class="form-select form-select-sm select2" id="filtroStatus" name="filtroStatus" data-placeholder="Todos">
                        <option value="">Todos</option>
                    </select>
                </div>

                <div class="col-12 mt-4 d-flex justify-content-between">
                    <div class="d-flex justify-content-end align-items-end gap-2">
                        <button onclick="limparFiltros()" type="button" class="btn btn-sm btn-outline-secondary" id="btnLimpar">
                            <i class="bi bi-x-lg me-1"></i> Limpar
                        </button>
                        <button onclick="carregarLotesPiloto()" type="button" class="btn btn-sm btn-primary" id="btnPesquisar">
                            <i class="bi bi-search me-1"></i> Pesquisar
                        </button>
                    </div>
                </div>
            </div>
        </div>

        <div class="card ">
            <div class="gap-2 p-3 border-bottom">
                <button id="btnNovoLotePiloto" class="btn btn-sm btn-primary" onclick="novoLotePiloto()">
                    <i class="bi bi-plus-lg"></i>
                    Novo
                </button>
            </div>

            <div class="table-wrapper">
                <div id="loadingOverlay" class="loading-overlay" style="display: none;">
                    <div class="spinner-border text-primary" role="status">
                        <span class="visually-hidden">Carregando...</span>
                    </div>
                </div>

                <table class="table table-hover mb-0" id="gradeLotePiloto">
                    <thead>
                        <tr>
                            <th>#</th>
                            <th>Planta</th>
                            <th>Data Inclusão</th>
                            <th>Data Status</th>
                            <th>Status</th>
                            <th>Cliente</th>
                            <th>Produto</th>
                            <th>Qtd. Paletes</th>
                            <th>Qtd. Pedidos</th>
                            <th>Motivo</th>
                            <th>Tipo LP</th>
                            <th>Usuário</th>
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

    <!-- Modal de detalhes do pedido -->
    <div class="modal fade" id="modalDetalhesPedido" tabindex="-1" aria-labelledby="modalDetalhesPedidoLabel" aria-hidden="true">
        <div class="modal-dialog modal-lg">
            <div class="modal-content">
                <div class="modal-header">
                    <h5 class="modal-title" id="modalDetalhesPedidoLabel">
                        Detalhes de Pedido Lote Piloto #<span id="idDetalhesPedido"></span>
                    </h5>
                    <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Fechar"></button>
                </div>
                <div class="modal-body">
                    <div class="card bg-light border-0 mb-3">
                        <div class="card-body py-2 px-3">
                            <div class="row g-2">
                                <div class="col-12 col-md-8">
                                    <small class="text-muted d-block">Cliente</small>
                                    <span class="fw-semibold" id="clienteDetalhesPedido"></span>
                                </div>
                                <div class="col-12 col-md-4">
                                    <small class="text-muted d-block">Produto</small>
                                    <span class="fw-semibold" id="produtoDetalhesPedido"></span>
                                </div>
                            </div>
                        </div>
                    </div>
                    <div id="conteudoDetalhesPedido"></div>
                </div>
            </div>
        </div>
    </div>

    <!-- Modal de Histórico de status -->
    <div class="modal fade" id="modalHistoricoStatus" tabindex="-1" aria-labelledby="modalHistoricoStatusLabel" aria-hidden="true">
        <div class="modal-dialog modal-xl">
            <div class="modal-content">
                <div class="modal-header">
                    <h5 class="modal-title" id="modalHistoricoStatusTitulo">
                        Histórico de alteração de status #<span id="idHistoricoStatus"></span>
                    </h5>
                    <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Fechar"></button>
                </div>
                <div class="modal-body">
                    <table class="table table-hover mb-0" id="gradeHistoricoStatus">
                        <thead>
                            <tr>
                                <th>#</th>
                                <th>Status Anterior</th>
                                <th>Status Posterior</th>
                                <th>Alterado Por</th>
                                <th>Quem Incluiu</th>
                                <th>Data Inclusão</th>
                            </tr>
                        </thead>
                        <tbody>
                            <tr>
                                <td colspan="6" class="text-center text-muted py-4">
                                    Nenhum registro encontrado
                                </td>
                            </tr>
                        </tbody>
                    </table>
                </div>
            </div>
        </div>
    </div>

    <!-- Modal de Histórico de justificativas -->
    <div class="modal fade" id="modalHistoricoJustificativas" tabindex="-1" aria-labelledby="modalHistoricoJustificativasLabel" aria-hidden="true">
        <div class="modal-dialog modal-xl">
            <div class="modal-content">
                <div class="modal-header">
                    <h5 class="modal-title" id="modalHistoricoJustificativasTitulo">
                        Histórico de alteração de justificativas #<span id="idHistoricoJustificativas"></span>
                    </h5>
                    <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Fechar"></button>
                </div>
                <div class="modal-body">
                    <table class="table table-hover mb-0" id="gradeHistoricoJustificativas">
                        <thead>
                            <tr>
                                <th>#</th>
                                <th>Tipo de Justificativa</th>
                                <th>Justificativa</th>
                                <th>Quem Incluiu</th>
                                <th>Data Inclusão</th>
                            </tr>
                        </thead>
                        <tbody>
                            <tr>
                                <td colspan="6" class="text-center text-muted py-4">
                                    Nenhum registro encontrado
                                </td>
                            </tr>
                        </tbody>
                    </table>
                </div>
            </div>
        </div>
    </div>

    <!-- Modal de exclusão -->
    <div class="modal fade" id="modalExclusao" tabindex="-1" aria-labelledby="modalExclusaoLabel" aria-hidden="true">
        <div class="modal-dialog modal-lg">
            <div class="modal-content">
                <div class="modal-header">
                        <h5 class="modal-title" id="modalExclusaoTitulo">
                        <span id="modalExclusaoTituloTexto"></span>
                    </h5>
                    <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Fechar"></button>
                </div>

                <div class="modal-body">
                    <input type="hidden" id="idExclusao" name="idExclusao">
                    <input type="hidden" id="statusExclusao" name="statusExclusao">
                    <input type="hidden" id="codigoProdutoExclusao" name="codigoProdutoExclusao">
                    <input type="hidden" id="codigoPlantaExclusao" name="codigoPlantaExclusao">

                    <div class="row mb-3">
                        <div class="col-12">
                            <label for="descricaoJustificativa" class="form-label"> Justificativa </label>
                            <textarea class="form-control form-control-sm" id="descricaoJustificativa" rows="5"></textarea>
                            <div class="invalid-feedback" id="feedbackJustificativa"></div>
                        </div>
                    </div>
                </div>

                <div class="modal-footer">
                    <button type="button" class="btn btn-sm btn-outline-secondary" data-bs-dismiss="modal">Cancelar</button>
                    <button onclick="excluirLotePiloto()" class="btn btn-sm btn-success" id="btnExcluir">
                        <i class="bi bi-check-lg me-1"></i>
                        <span>Salvar</span>
                    </button>
                </div>
            </div>
        </div>
    </div>
</body>
</html>