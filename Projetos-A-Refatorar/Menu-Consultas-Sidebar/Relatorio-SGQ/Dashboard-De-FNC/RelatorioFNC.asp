<!DOCTYPE html>
<html lang="pt-br">
<head>
    <title>Smurfit Kappa - Relatório de FNC</title>
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
            max-width: 100%;
        }

        .form-label {
            font-weight: 600;
        }

        .required-field {
            color: #dc3545;
        }

        .table-wrapper {
            overflow: auto;
            max-height: 70vh;
        }

        #gradeFNC {
            width: max-content;
            min-width: 100%;
        }

        #gradeFNC th,
        #gradeFNC td {
            white-space: nowrap;
            vertical-align: top;
        }

        #gradeFNC td.multi-line {
            white-space: normal;
            min-width: 120px;
        }

        #gradeFNC th:first-child,
        #gradeFNC td:first-child {
            position: sticky;
            left: 0;
            z-index: 5;
            white-space: nowrap;
        }
        
        #gradeFNC td:first-child {
            background: linear-gradient(to left, rgba(0,0,0,0.1) 0%, transparent 1px),#fff;
            background-origin: border-box;
            background-clip: padding-box;
        }
        
        #gradeFNC thead th:first-child {
            background: linear-gradient(to left, rgba(0,0,0,0.1) 0%, transparent 1px), #f5f5f5;
            background-origin: border-box;
            background-clip: padding-box;
            z-index: 10;
        }

        #gradeFNC th:last-child,
        #gradeFNC td:last-child {
            position: sticky;
            right: 0;
            z-index: 5;
            white-space: nowrap;
        }
        
        #gradeFNC td:last-child {
            background: linear-gradient(to right, rgba(0,0,0,0.1) 0%, transparent 1px),#fff;
            background-origin: border-box;
            background-clip: padding-box;
        }
        
        #gradeFNC thead th:last-child {
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

        .no-records {
            padding: 40px;
            text-align: center;
            background: white;
            color: #6c757d;
        }

        .truncate-text {
            max-width: 250px;
            white-space: nowrap;
            overflow: hidden;
            text-overflow: ellipsis;
            cursor: help;
        }

        .z-index-9999 {
            z-index: 9999 !important;
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

        .badge-aberta {
            background-color: #d1fae5;
            color: #065f46;
        }

        .badge-respondida {
            background-color: #dbeafe;
            color: #1e40af;
        }

        .badge-encerrada {
            background-color: #fce7f3;
            color: #9f1239;
        }

        .badge-status-aprovada {
            background-color: #d1fae5;
            color: #065f46;
        }

        .badge-status-nao-aprovada {
            background-color: #fee2e2;
            color: #991b1b;
        }

        .badge-validacao-registrada {
            background-color: #e0e7ff;
            color: #3730a3;
        }

        .badge-validacao-em-analise {
            background-color: #fef3c7;
            color: #92400e;
        }

        .badge-validacao-procede {
            background-color: #d1fae5;
            color: #065f46;
        }

        .badge-validacao-nao-procede {
            background-color: #fee2e2;
            color: #991b1b;
        }

        .badge-validacao-cancelada {
            background-color: #f3f4f6;
            color: #4b5563;
        }

        .item-list {
            display: flex;
            flex-direction: column;
            gap: 2px;
        }

        .item-list span {
            display: block;
        }
    </style>
    
    <script src="https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/jspdf-autotable/3.5.31/jspdf.plugin.autotable.min.js"></script>
    <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.0.2/dist/js/bootstrap.bundle.min.js" integrity="sha384-MrcW6ZMFYlzcLA8Nl+NtUVF0sA7MsXsP1UyJoMp4YLEuNSfAP+JcXn/tWtIaxVXM" crossorigin="anonymous"></script>
    <script src="https://cdn.jsdelivr.net/npm/select2@4.1.0-rc.0/dist/js/select2.min.js"></script>
    <script src="../../../components/toast.js?v=<%=Now()%>"></script>
    <script src="/corporativo/funcoes/funcoesJAVASCRIPT.js?v=<%=Now()%>"></script>
    <script src="/corporativo/components/dropdownMenu.js?v=<%=Now()%>"></script>

    <script>
        let dadosFNC = [];
        let carregando = false;
        let limitesNQA = {};

        function inicializarSelect2() {
            $('.select2').select2({
                theme: 'bootstrap-5',
                width: '100%',
                dropdownCssClass: "z-index-9999",
                allowClear: true,
                placeholder: function() {
                    return $(this).data('placeholder');
                }
            });
        }

        function carregarSetores() {
            const $select = $('#filtroSetor');
            $select.empty().append('<option value="">Todos os setores</option>');

            const params = {
                acao: 'listarSetores'
            };

            $.ajax({
                url: 'RelatorioFNCRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: params,
                dataType: 'json',
                success: function(response) {
                    if (!response.sucesso) {
                        console.error("Erro ao carregar setores: " + response.mensagem);
                        return;
                    }

                    response.dados.forEach(function(item) {
                        $select.append($('<option>', { value: item.id, text: item.descricao }));
                    });
                },
                error: function(xhr, status, error) {
                    console.error('Erro ao carregar setores:', error);
                }
            });
        }

        function carregarDefeitos() {
            const $select = $('#filtroDefeito');
            $select.empty().append('<option value="">Todos os defeitos</option>');

            const params = {
                acao: 'listarDefeitos',
                filtroTipoDefeito: $('#filtroTipoDefeito').val()
            };

            $.ajax({
                url: 'RelatorioFNCRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: params,
                dataType: 'json',
                success: function(response) {
                    if (!response.sucesso) {
                        console.error("Erro ao carregar defeitos: " + response.mensagem);
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
            const $select = $('#filtroDisposicao');
            $select.empty().append('<option value="">Todas as disposições</option>');

            const params = {
                acao: 'listarDisposicoes'
            };

            $.ajax({
                url: 'RelatorioFNCRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: params,
                dataType: 'json',
                success: function(response) {
                    if (!response.sucesso) {
                        console.error("Erro ao carregar disposições: " + response.mensagem);
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

        function carregarResponsaveis() {
            const $select = $('#filtroResponsavel');
            $select.empty().append('<option value="">Todos os responsáveis</option>');

            const params = {
                acao: 'listarResponsaveis',
                filtroFilial: $('#filtroFilial').val()
            };

            $.ajax({
                url: 'RelatorioFNCRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: params,
                dataType: 'json',
                success: function(response) {
                    if (!response.sucesso) {
                        console.error("Erro ao carregar responsáveis: " + response.mensagem);
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

        function exibirEspelhoCompleto(idFNC) {
            const baseUrl = `${location.protocol}//${location.host}`;
            const url = `${baseUrl}/corporativo/sgq/fnc/espelho.asp?id=${idFNC}&mostrarTodasRespostas=1`;
            window.open(url, '_blank');
        }
                
        function abrirModalRespostas(idFNC) {  
            var modal = new bootstrap.Modal(document.getElementById('modalRespostas'));
            modal.show();
            carregarRespostas(idFNC);
        }

        function carregarRespostas(idFNC) {
            $.ajax({
                url: '../FNCRespostas.asp?v=' + Math.random(),
                type: 'GET',
                data: { idFNC: idFNC },
                success: function(response) {
                    $('#modalRespostasBody').html(response);
                },
                error: function(xhr, status, error) {
                    $('#modalRespostasBody').html('<div class="alert alert-danger">Erro ao carregar histórico: ' + error + '</div>');
                }
            });
        }

        function abrirModalObservacoes(idFNC) {  
            var modal = new bootstrap.Modal(document.getElementById('modalObservacoes'));
            modal.show();
            carregarObservacoes(idFNC);
        }

        function visualizarFNC(id) {
            window.location.href = `../Principal/FNCFormularioEdicao.asp?id=${id}&visualizar=1`;
        }

        function carregarObservacoes(idFNC) {
            $.ajax({
                url: '../FNCObservacoes.asp?v=' + Math.random(),
                type: 'GET',
                data: { idFNC: idFNC },
                success: function(response) {
                    $('#modalObservacoesBody').html(response);
                },
                error: function(xhr, status, error) {
                    $('#modalObservacoesBody').html('<div class="alert alert-danger">Erro ao carregar histórico: ' + error + '</div>');
                }
            });
        }

        function pesquisarCliente() {
            const codPlanta = $('#filtroFilial').val();
            const baseUrl = `${location.protocol}//${location.host}`;
            const url = `${baseUrl}/corporativo/buscaCLIENTE.asp?id1=filtroCodCliente&id2=filtroCodLoja&id3=nomeCliente&id6=${codPlanta}`;
            window.open(url, '', "width=650,height=400");
        }

        function validarFiltros() {
            if (
                (!$('#filtroDataInclusaoDe').val() || !$('#filtroDataInclusaoAte').val())
                && 
                (!$('#filtroDataEnvioClienteDe').val() || !$('#filtroDataEnvioClienteAte').val())) 
            {
                Toast.error("Preencha a data de inclusão ou envio ao cliente");
                return false;
            }
            return true;
        }

        function exibirLoading(exibir) {
            carregando = exibir;
            if (exibir) {
                $('#loadingOverlay').show();
            } else {
                $('#loadingOverlay').hide();
            }
        }

        function pesquisar() {
            if (!validarFiltros()) return;
            
            exibirLoading(true);

            const params = {
                acao: 'listarFNCs',
                filtroFilial: $('#filtroFilial').val(),
                filtroDataInclusaoDe: $('#filtroDataInclusaoDe').val(),
                filtroDataInclusaoAte: $('#filtroDataInclusaoAte').val(),
                filtroDataEnvioClienteDe: $('#filtroDataEnvioClienteDe').val(),
                filtroDataEnvioClienteAte: $('#filtroDataEnvioClienteAte').val(),
                filtroIdFNC: $('#filtroIdFNC').val(),
                filtroPedido: $('#filtroPedido').val(),
                filtroNotaFiscal: $('#filtroNotaFiscal').val(),
                filtroCodCliente: $('#filtroCodCliente').val(),
                filtroCodLoja: $('#filtroCodLoja').val(),
                filtroSetor: $('#filtroSetor').val()?.join(','),
                filtroSituacao: $('#filtroSituacao').val()?.join(','),
                filtroValidacao: $('#filtroValidacao').val()?.join(','),
                filtroDefeito: $('#filtroDefeito').val()?.join(','),
                filtroTipoDefeito: $('#filtroTipoDefeito').val()?.join(','),
                filtroDisposicao: $('#filtroDisposicao').val()?.join(','),
                filtroResponsavel: $('#filtroResponsavel').val()?.join(','),
            };

            $.ajax({
                url: 'RelatorioFNCRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: params,
                dataType: 'json',
                success: function(response) {
                    exibirLoading(false);

                    if (!response.sucesso) {
                        alert('Erro ao pesquisar: ' + response.mensagem);
                        return;
                    }

                    dadosFNC = response.dados || [];
                    renderizarGrade();
                    controlarExibicaoAviso();
                },
                error: function(xhr, status, error) {
                    exibirLoading(false);
                    console.error('Erro ao pesquisar:', error);
                    alert('Erro ao realizar a pesquisa');
                }
            });
        }

        function formatarCampoMultiplo(valor) {
            if (!valor || valor === '-') 
                return '-';
            
            const itens = valor.split('|').filter(item => item.trim() !== '');

            if (itens.length === 0) return '-';

            return '<div class="item-list">' + itens.map(item => '<span>' + item.trim() + '</span>').join('') + '</div>';
        }

        function renderizarGrade() {
            const $tbody = $('#gradeFNC tbody');
            
            $tbody.empty();
            
            if (dadosFNC.length === 0) {
                $tbody.html(`
                    <tr>
                        <td colspan="28" class="text-center text-muted py-4">
                            Utilize os filtros acima para pesquisar
                        </td>
                    </tr>
                `);
                return;
            }

            dadosFNC.forEach(function(item) {
                const situacaoBadge = obterBadgeSituacao(item.situacao);
                const statusBadge = obterBadgeStatus(item.status);
                const validacaoBadge = obterBadgeValidacao(item.validacao);
                const baseUrl = `${location.protocol}//${location.host}`;
                const urlFNC = `${baseUrl}/corporativo/sgq/fnc/espelho.asp?id=${item.idFNC}&filial=${item.codFilial}`;
                
                if(item.tipoDefeito != 'PRO') {
                    statusNQA = '-'
                } else if(parseFloat(item.coeficienteNQA) < parseFloat(limitesNQA.critico)) {
                    statusNQA = `${formatarMoeda(item.coeficienteNQA)} | APROVADO`
                } else if ( (parseFloat(item.coeficienteNQA) >= parseFloat(limitesNQA.critico)) && (parseFloat(item.coeficienteNQA) < parseFloat(limitesNQA.grave)) ) {
                    statusNQA = `${formatarMoeda(item.coeficienteNQA)} | CRITICO`
                } else {
                    statusNQA = `${formatarMoeda(item.coeficienteNQA)} | GRAVE`
                }

                $tbody.append(`
                    <tr>
                        <td>
                            <a href="${urlFNC}" target="_blank">
                                <i class="bi bi-file-earmark-text"></i>
                                ${item.idFNC}
                            </a>
                        </td>
                        <td>${item.nomeFilial || '-'}</td>
                        <td>${item.dataInclusao || '-'}</td>
                        <td>${item.dataProducao || '-'}</td>
                        <td>${item.dataOcorrencia || '-'}</td>
                        <td>${item.dataEnvioCliente || '-'}</td>
                        <td>${item.dataConclusaoPrevista || '-'}</td>
                        <td>${item.quantidade || '-'}</td>
                        <td>${situacaoBadge}</td>
                        <td>${statusBadge}</td>
                        <td>${validacaoBadge}</td>
                        <td>${item.codCliente || '-'}</td>
                        <td>${item.codLoja || '-'}</td>
                        <td>${item.nomeCliente || '-'}</td>
                        <td>${item.cidadeEstado || '-'}</td>
                        <td class="multi-line">${formatarCampoMultiplo(item.codProduto)}</td>
                        <td class="multi-line">${formatarCampoMultiplo(item.referencia)}</td>
                        <td class="multi-line">${formatarCampoMultiplo(item.grupoProduto)}</td>
                        <td>${item.maquina || '-'}</td>
                        <td>${item.turno || '-'}</td>
                        <td>${item.setor || '-'}</td>
                        <td>${item.descricaoDefeito || '-'}</td>
                        <td>${obterDescricaoTipoDefeito(item.tipoDefeito || '-')}</td>
                        <td>${item.disposicao || '-'}</td>
                        <td>
                            <div class="truncate-text" title="${item.descricaoDisposicao || ''}">
                                ${item.descricaoDisposicao || '-'}
                            </div>
                        </td>
                        <td>
                            <div class="truncate-text" title="${item.eficacia || ''}">
                                ${item.eficacia || '-'}
                            </div>
                        </td>
                        <td class="multi-line">${formatarCampoMultiplo(item.notasFiscais)}</td>
                        <td>${statusNQA}</td>
                        <td>${item.valorAgregado || '-'}</td>
                        <td>${item.valorReclamado || '-'}</td>
                        <td class="multi-line">${formatarCampoMultiplo(item.precoUnitario)}</td>
                        <td class="multi-line">${formatarCampoMultiplo(item.quantidadePedidos)}</td>
                        <td class="multi-line">${formatarCampoMultiplo(item.valorPedido)}</td>
                        <td class="multi-line">${formatarCampoMultiplo(item.dataFaturamento)}</td>
                        <td>${item.nomeAnalista || '-'}</td>
                        <td>${item.nomeResponsavel || '-'}</td>
                        <td>${item.diasAteResposta || 0}</td>
                        <td>${item.diasAteEncerramento || '-'}</td>
                        <td style="text-align: center">
                            <dropdown-menu label="Ações">
                                <a class="dropdown-item" onclick="visualizarFNC('${item.idFNC}')"><i class="bi bi-eye me-1"></i> Detalhes</a>
                                <a class="dropdown-item" onclick="abrirModalRespostas('${item.idFNC}')"><i class="bi bi bi-reply me-1"></i> Respostas</a>
                                <a class="dropdown-item" onclick="abrirModalObservacoes('${item.idFNC}')"><i class="bi bi-chat-dots me-1"></i> Observações</a>
                                <a class="dropdown-item" onclick="exibirEspelhoCompleto('${item.idFNC}')"><i class="bi bi-file-earmark-text"></i> Espelho (Todas respostas)</a>
                           </dropdow-menu>
                        </td>
                    </tr>
                `);
            });
        }

        function obterDescricaoSituacao(situacao) {
            switch(situacao) {
                case '1':
                    return 'ABERTA';
                case '2':
                    return 'RESPONDIDA';
                case '3':
                    return 'ENCERRADA';
                default:
                    return '';
            }
        }

        function obterDescricaoTipoDefeito(defeito) {
            switch(defeito) {
                case 'ADM':
                    return 'ADMINISTRATIVO';
                case 'PRO':
                    return 'INDUSTRIAL';
                default:
                    return '';
            }
        }

        function obterDescricaoStatus(status) {
            switch(status) {
                case '1':
                    return 'APROVADA';
                case '0':
                    return 'NÃO APROVADA';
                default:
                    return '';
            }
        }

        function obterDescricaoValidacao(validacao) {
            switch(validacao) {
                case '1':
                    return 'REGISTRADA';
                case '2':
                    return 'EM ANÁLISE';
                case '3':
                    return 'PROCEDE';
                case '4':
                    return 'NÃO PROCEDE';
                case '5':
                    return 'CANCELADA';
                default:
                    return '';
            }
        }

        function obterBadgeSituacao(situacao) {
            switch(situacao) {
                case '1':
                    return '<span class="badge badge-aberta">ABERTA</span>';
                case '2':
                    return '<span class="badge badge-respondida">RESPONDIDA</span>';
                case '3':
                    return '<span class="badge badge-encerrada">ENCERRADA</span>';
                default:
                    return '';
            }
        }

        function obterBadgeStatus(status) {
            switch(status) {
                case '1':
                    return '<span class="badge badge-status-aprovada">APROVADA</span>';
                case '0':
                    return '<span class="badge badge-status-nao-aprovada">NÃO APROVADA</span>';
                default:
                    return '-';
            }
        }

        function obterBadgeValidacao(validacao) {
            switch(validacao) {
                case '1':
                    return '<span class="badge badge-validacao-registrada">REGISTRADA</span>';
                case '2':
                    return '<span class="badge badge-validacao-em-analise">EM ANÁLISE</span>';
                case '3':
                    return '<span class="badge badge-validacao-procede">PROCEDE</span>';
                case '4':
                    return '<span class="badge badge-validacao-nao-procede">NÃO PROCEDE</span>';
                case '5':
                    return '<span class="badge badge-validacao-cancelada">CANCELADA</span>';
                default:
                    return '-';
            }
        }

        function limparFiltros() {
            $('#filtroFilial').val('').trigger('change');
            $('#filtroSetor').val('').trigger('change');
            $('#filtroIdFNC').val('');
            $('#filtroPedido').val('');
            $('#filtroNotaFiscal').val('');
            $('#filtroSituacao').val('').trigger('change');
            $('#filtroValidacao').val('').trigger('change');
            $('#filtroDataInclusaoDe').val('');
            $('#filtroDataInclusaoAte').val('');
            $('#filtroDataEnvioClienteDe').val('');
            $('#filtroDataEnvioClienteAte').val('');
            $('#filtroDefeito').val('').trigger('change');
            $('#filtroTipoDefeito').val('').trigger('change');
            $('#filtroDisposicao').val('').trigger('change');
            $('#filtroResponsavel').val('').trigger('change');
            $('#filtroCodCliente').val('');
            $('#filtroCodLoja').val('');
            
            dadosFNC = [];
            renderizarGrade();
        }

        function formatarCampoMultiploExport(valor, separador = '\n') {
            if (!valor || valor === '-') return '-';
            const itens = valor.split('|').filter(item => item.trim() !== '');
            if (itens.length === 0) return '-';
            return itens.map(item => item.trim()).join(separador);
        }

        function exportarPDF() {
            if (dadosFNC.length === 0) {
                alert('Não há dados para exportar');
                return;
            }

            const { jsPDF } = window.jspdf;

            var doc = new jsPDF({
                orientation: 'landscape', 
                unit: 'mm',
                format: 'a0'
            });

            const now = new Date();

            const options = {
                day: '2-digit',
                month: '2-digit',
                year: 'numeric',
                hour: '2-digit',
                minute: '2-digit',
                second: '2-digit',
                hour12: false 
            };

            const dataAtual = new Intl.DateTimeFormat('pt-BR', options).format(now);

            const colunas = [
                'FNC', 
                'Nome Filial',
                'Data Inclusão', 
                'Data de Produção',
                'Data da Ocorrência', 
                'Data Envio ao Cliente', 
                'Data de Conclusão Prevista',
                'Qtde.', 
                'Situação',
                'Status', 
                'Validação',
                'Cód. Cliente', 
                'Cód. Loja', 
                'Razão Social', 
                'Cidade/Estado', 
                'Cód. Produto', 
                'Referência', 
                'Grupo Produto',
                'Máquina', 
                'Turno', 
                'Setor',
                'Defeito',
                'Tipo Defeito',
                'Disposição',
                'Descrição Disposição',
                'Eficácia',
                'Nota Fiscal SW',
                'Coeficiente NQA',
                'Valor Agregado',
                'Valor Reclamado',
                'Preço Unitário',
                'Qtde. Vendida',
                'Total NF',
                'Data Faturamento',
                'Analista',
                'Responsável',
                'Dias Aberto Resposta',
                'Dias Aberto Encerramento'
            ];

            const dados = dadosFNC.map(item => [
                item.idFNC || '-',
                item.nomeFilial || '-',
                item.dataInclusao || '-',
                item.dataProducao || '-',
                item.dataOcorrencia || '-',
                item.dataEnvioCliente || '-',
                item.dataConclusaoPrevista || '-',
                item.quantidade || '-',
                obterDescricaoSituacao(item.situacao) || '-',
                obterDescricaoStatus(item.status || '-'),
                obterDescricaoValidacao(item.validacao || '-'),
                item.codCliente || '-',
                item.codLoja || '-',
                item.nomeCliente || '-',
                item.cidadeEstado || '-',
                formatarCampoMultiploExport(item.codProduto, ', '),
                formatarCampoMultiploExport(item.referencia, ', '),
                formatarCampoMultiploExport(item.grupoProduto, ', '),
                item.maquina || '-',
                item.turno || '-',
                item.setor || '-',
                item.descricaoDefeito || '-',
                obterDescricaoTipoDefeito(item.tipoDefeito || '-'),
                item.disposicao || '-',
                item.descricaoDisposicao || '-',
                item.eficacia || '-',
                formatarCampoMultiploExport(item.notasFiscais, ', '),
                item.tipoDefeito != 'PRO' ? '-' : formatarMoeda(item.coeficienteNQA),
                item.valorAgregado || '-',
                item.valorReclamado || '-',
                formatarCampoMultiploExport(item.precoUnitario, ', '),
                formatarCampoMultiploExport(item.quantidadePedidos, ', '),
                formatarCampoMultiploExport(item.valorPedido, ', '),
                formatarCampoMultiploExport(item.dataFaturamento, ', '),
                item.nomeAnalista || '-',
                item.nomeResponsavel || '-',
                item.diasAteResposta || 0,
                item.diasAteEncerramento || '-'
            ]);

            var yPosition = 15;
            let filtrosAplicados = [];
                
            const filtroDataInclusaoDe = $('#filtroDataInclusaoDe').val();
            const filtroDataInclusaoAte = $('#filtroDataInclusaoAte').val();
            const filtroEnvioClienteDe = $('#filtroEnvioClienteDe').val();
            const filtroEnvioClienteAte = $('#filtroEnvioClienteAte').val();
            const filtroFilial = $('#filtroFilial').val() && $('#filtroFilial option:selected').text();
            const filtroCodCliente = $('#filtroCodCliente').val();
            const filtroCodLoja = $('#filtroCodLoja').val();
            const filtroIdFNC = $('#filtroIdFNC').val();
            const filtroPedido = $('#filtroPedido').val();
            const filtroNotaFiscal = $('#filtroNotaFiscal').val();
           
            const filtroSetor = (data = $('#filtroSetor').select2('data')).length && data.map(item => item.text);
            const filtroSituacao = (data = $('#filtroSituacao').select2('data')).length && data.map(item => item.text);
            const filtroValidacao = (data = $('#filtroValidacao').select2('data')).length && data.map(item => item.text);
            const filtroDefeito = (data = $('#filtroDefeito').select2('data')).length && data.map(item => item.text);
            const filtroTipoDefeito = (data = $('#filtroTipoDefeito').select2('data')).length && data.map(item => item.text);
            const filtroDisposicao = (data = $('#filtroDisposicao').select2('data')).length && data.map(item => item.text);
            const filtroResponsavel = (data = $('#filtroResponsavel').select2('data')).length && data.map(item => item.text);

            if (filtroDataInclusaoDe) 
                filtrosAplicados.push('Data Inclusão (De): ' + new Date(`${filtroDataInclusaoDe} 03:00`).toLocaleDateString('pt-BR'));

            if (filtroDataInclusaoAte) 
                filtrosAplicados.push('Data Inclusão (Até): ' + new Date(`${filtroDataInclusaoAte} 03:00`).toLocaleDateString('pt-BR') );

            if (filtroDataEnvioClienteDe) 
                filtrosAplicados.push('Data Envio Cliente (De): ' + new Date(`${filtroDataEnvioClienteDe} 03:00`).toLocaleDateString('pt-BR'));

            if (filtroDataEnvioClienteAte) 
                filtrosAplicados.push('Data Envio Cliente (Até): ' + new Date(`${filtroDataEnvioClienteAte} 03:00`).toLocaleDateString('pt-BR') );

            if (filtroFilial) 
                filtrosAplicados.push('Filial: ' + filtroFilial);

            if (filtroCodCliente) 
                filtrosAplicados.push('Cód. Cliente: ' + filtroCodCliente);

            if (filtroCodLoja) 
                filtrosAplicados.push('Cód. Loja: ' + filtroCodLoja);

            if (filtroIdFNC) 
                filtrosAplicados.push('Id FNC: ' + filtroIdFNC);

            if (filtroPedido) 
                filtrosAplicados.push('Pedido: ' + filtroPedido);

            if (filtroNotaFiscal) 
                filtrosAplicados.push('Nota Fiscal: ' + filtroNotaFiscal);

            if (filtroSetor && filtroSetor.length > 0) 
                filtrosAplicados.push('Setor: ' + filtroSetor.join(', '));

            if (filtroSituacao && filtroSituacao.length > 0) 
                filtrosAplicados.push('Situação: ' + filtroSituacao.join(', '));

            if (filtroValidacao && filtroValidacao.length > 0) 
                filtrosAplicados.push('Validação: ' + filtroValidacao.join(', '));

            if (filtroDefeito && filtroDefeito.length > 0) 
                filtrosAplicados.push('Defeito: ' + filtroDefeito.join(', '));

            if (filtroTipoDefeito && filtroTipoDefeito.length > 0) 
                filtrosAplicados.push('Tipo Defeito: ' +  filtroTipoDefeito.join(', '));

            if (filtroDisposicao && filtroDisposicao.length > 0) 
                filtrosAplicados.push('Disposição: ' + filtroDisposicao.join(', '));

            if (filtroResponsavel && filtroResponsavel.length > 0) 
                filtrosAplicados.push('Responsável: ' + filtroResponsavel.join(', '));

            if (filtrosAplicados.length > 0) {
                doc.text('Filtros Aplicados:', 14, yPosition);
                yPosition += 7;
                doc.setFont(undefined, 'normal');

                for (var i = 0; i < filtrosAplicados.length; i++) {
                    doc.text('  • ' + filtrosAplicados[i], 14, yPosition);
                    yPosition += 7;
                }

                yPosition += 2;
            }

            doc.setFontSize(16);
            doc.text('Relatório de Respostas de FNC', 14, yPosition);
            yPosition += 7;

            doc.setFontSize(10);
            doc.text(`Data: ${dataAtual} | ${dados.length} registro(s) encontrado(s)`, 14, yPosition);
            yPosition += 7;

            doc.autoTable({
                head: [colunas],
                body: dados,
                startY: yPosition,
                theme: 'striped',
                styles: { 
                    fontSize: 8,
                    cellPadding: 2
                },
                headStyles: { 
                    fillColor: [211, 211, 212],
                    textColor: [0, 0, 0],
                    fontStyle: 'bold'
                },
                columnStyles: {
                    0: { cellWidth: 25 },
                    1: { cellWidth: 'auto' },
                    2: { cellWidth: 50 },
                    3: { cellWidth: 50 }
                },
                margin: { top: 28 }
            });
                            
            window.open(doc.output('bloburl'), '_blank');
        }


        function exportarExcel() {
            if (dadosFNC.length === 0) {
                alert('Não há dados para exportar');
                return;
            }

            const nomeArquivo = 'Relatorio_FNC.csv';
            const colunas = [
                'FNC', 
                'Nome Filial',
                'Data Inclusão', 
                'Data Produção', 
                'Data da Ocorrência', 
                'Data Envio ao Cliente',
                'Data de Conclusão Prevista',
                'Qtde.', 
                'Situação',
                'Status', 
                'Validação',
                'Cód. Cliente', 
                'Cód. Loja', 
                'Razão Social',
                'Cidade/Estado',  
                'Cód. Produto', 
                'Referência', 
                'Grupo Produto',
                'Máquina', 
                'Turno', 
                'Setor',
                'Defeito',
                'Tipo Defeito',
                'Disposição',
                'Descrição Disposição',
                'Eficácia',
                'Nota Fiscal SW',
                'Coeficiente NQA',
                'Valor Agregado',
                'Valor Reclamado',
                'Preço Unitário',
                'Qtde. Vendida',
                'Total NF',
                'Data Faturamento',
                'Analista',
                'Responsável',
                'Dias Aberto Resposta',
                'Dias Aberto Encerramento'
            ];

            let csvContent = '\uFEFF'; 
            csvContent += colunas.join(';') + '\r\n';

            dadosFNC.forEach(item => {
                const linha = [
                    item.idFNC || '-',
                    item.nomeFilial || '-',
                    item.dataInclusao || '-',
                    item.dataProducao || '-',
                    item.dataOcorrencia || '-',
                    item.dataEnvioCliente || '-',
                    item.dataConclusaoPrevista || '-',
                    item.quantidade || '-',
                    obterDescricaoSituacao(item.situacao) || '-',
                    obterDescricaoStatus(item.status || '-'),
                    obterDescricaoValidacao(item.validacao || '-'),
                    item.codCliente || '-',
                    item.codLoja || '-',
                    escapeCSV(item.nomeCliente || '-'),
                    escapeCSV(item.cidadeEstado || '-'),
                    escapeCSV(formatarCampoMultiploExport(item.codProduto, ', ')),
                    escapeCSV(formatarCampoMultiploExport(item.referencia, ', ')),
                    escapeCSV(formatarCampoMultiploExport(item.grupoProduto, ', ')),
                    item.maquina || '-',
                    item.turno || '-',
                    item.setor || '-',
                    escapeCSV(item.descricaoDefeito || '-'),
                    obterDescricaoTipoDefeito(item.tipoDefeito || '-'),
                    escapeCSV(item.disposicao || '-'),
                    escapeCSV(item.descricaoDisposicao || '-'),
                    escapeCSV(item.eficacia || '-'),
                    escapeCSV(formatarCampoMultiploExport(item.notasFiscais, ', ')),
                    item.tipoDefeito != 'PRO' ? '-' : formatarMoeda(item.coeficienteNQA),
                    item.valorAgregado || '-',
                    item.valorReclamado || '-',
                    escapeCSV(formatarCampoMultiploExport(item.precoUnitario, ', ')),
                    escapeCSV(formatarCampoMultiploExport(item.quantidadePedidos, ', ')),
                    escapeCSV(formatarCampoMultiploExport(item.valorPedido, ', ')),
                    escapeCSV(formatarCampoMultiploExport(item.dataFaturamento, ', ')),
                    item.nomeAnalista || '-',
                    item.nomeResponsavel || '-',
                    item.diasAteResposta || 0,
                    item.diasAteEncerramento || '-'
                ];
                csvContent += linha.join(';') + '\r\n';
            });

            const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
            const link = document.createElement('a');
            
            if (navigator.msSaveBlob) { 
                navigator.msSaveBlob(blob, nomeArquivo);
            } else {
                const url = URL.createObjectURL(blob);
                link.href = url;
                link.download = nomeArquivo;
                link.style.display = 'none';
                document.body.appendChild(link);
                link.click();
                document.body.removeChild(link);
                URL.revokeObjectURL(url);
            }
        }

        function escapeCSV(valor) {
            if (!valor) return '""';
            valor = valor.replace(/[\r\n]+/g, ' ').trim();
            valor = valor.replace(/"/g, '""');
            return '"' + valor + '"';
        }
    
        function anexarEventos() {
            $('#btnPesquisar').on('click', pesquisar);
            $('#btnLimpar').on('click', limparFiltros);
            $('#btnExportarExcel').on('click', exportarExcel);
            $('#btnExportarPDF').on('click', exportarPDF);
            $('#filtroFilial').on('change', carregarResponsaveis);
            $('#filtroTipoDefeito').on('change', carregarDefeitos);
        }

        function controlarExibicaoAviso() {    
            if (!$('#filtroFilial').val()) {
                $('#avisoFilial').removeClass('d-none');
            } else {
                $('#avisoFilial').addClass('d-none');
            }
        }

        function carregarLimitesNQA() {
            return new Promise((resolve, reject) => {
                $.ajax({
                    url: '../Principal/FNCRepositorio.asp?v=<%=Now()%>',
                    type: 'GET',
                    data: {
                        acao: 'obterLimitesNQA'
                    },
                    dataType: 'json',
                    success: function(response) {
                        resolve(response); 
                    },
                    error: function(xhr, status, error) {
                        reject(error); 
                    }
                });
            });
        }

        $(document).ready(async function() {
            respostaNQA = await carregarLimitesNQA();
            if (respostaNQA.sucesso) limitesNQA = respostaNQA.dados;

            inicializarSelect2();
            carregarSetores();
            carregarDefeitos();
            carregarDisposicoes();
            carregarResponsaveis();
            anexarEventos();
            renderizarGrade();

            const urlParams = new URLSearchParams(window.location.search);
            if (urlParams.size === 0) return;

            setTimeout(() => {
                if (urlParams.has('filtroDataInclusaoDe')) 
                    $('#filtroDataInclusaoDe').val(urlParams.get('filtroDataInclusaoDe'));

                if (urlParams.has('filtroDataInclusaoAte')) 
                    $('#filtroDataInclusaoAte').val(urlParams.get('filtroDataInclusaoAte'));

                if (urlParams.has('filtroDataEnvioClienteDe')) 
                    $('#filtroDataEnvioClienteDe').val(urlParams.get('filtroDataEnvioClienteDe'));

                if (urlParams.has('filtroDataEnvioClienteAte')) 
                    $('#filtroDataEnvioClienteAte').val(urlParams.get('filtroDataEnvioClienteAte'));

                if (urlParams.has('filtroFilial')) 
                    $('#filtroFilial').val(urlParams.get('filtroFilial')).trigger('change');

                if (urlParams.has('filtroCodCliente')) 
                    $('#filtroCodCliente').val(urlParams.get('filtroCodCliente'));

                if (urlParams.has('filtroCodLoja')) 
                    $('#filtroCodLoja').val(urlParams.get('filtroCodLoja'));

                if (urlParams.has('filtroSetor')) 
                    $('#filtroSetor').val(urlParams.get('filtroSetor').split(',')).trigger('change');

                if (urlParams.has('filtroSituacao')) 
                    $('#filtroSituacao').val(urlParams.get('filtroSituacao').split(',')).trigger('change');

                if (urlParams.has('filtroValidacao')) 
                    $('#filtroValidacao').val(urlParams.get('filtroValidacao').split(',')).trigger('change');
                
                if (urlParams.has('filtroDefeito')) 
                    $('#filtroDefeito').val(urlParams.get('filtroDefeito').split(',')).trigger('change');

                pesquisar();
            }, 1000)
        });
    </script>
</head>
<body>
    <div class="container my-4">
        <h3 class="mb-4">Relatório de FNC</h3>

        <div class="card p-3 mb-4">
            <div class="row g-3">
                <div class="col-12 col-md-4 col-lg-3">
                    <label for="filtroDataInclusaoDe" class="form-label">Data Inclusão De</label>
                    <input type="date" class="form-control form-control-sm" id="filtroDataInclusaoDe" name="filtroDataInclusaoDe">
                </div>

                <div class="col-12 col-md-4 col-lg-3">
                    <label for="filtroDataInclusaoAte" class="form-label">Data Inclusão Até</label>
                    <input type="date" class="form-control form-control-sm" id="filtroDataInclusaoAte" name="filtroDataInclusaoAte">
                </div>

                <div class="col-12 col-md-4 col-lg-3">
                    <label for="filtroDataEnvioClienteDe" class="form-label">Data Envio Cliente De</label>
                    <input type="date" class="form-control form-control-sm" id="filtroDataEnvioClienteDe" name="filtroDataEnvioClienteDe">
                </div>

                <div class="col-12 col-md-4 col-lg-3">
                    <label for="filtroDataEnvioClienteAte" class="form-label">Data Envio Cliente Até</label>
                    <input type="date" class="form-control form-control-sm" id="filtroDataEnvioClienteAte" name="filtroDataEnvioClienteAte">
                </div>
                
                <div class="col-12 col-md-4 col-lg-3">
                    <label for="filtroFilial" class="form-label">Filial</label>
                    <select class="form-select form-select-sm select2" id="filtroFilial" name="filtroFilial" data-placeholder="Todas as filiais">
                        <option value="">Todas as filiais</option>
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
                    <label for="filtroIdFNC" class="form-label">Id FNC</label>
                    <input type="number" class="form-control form-control-sm" id="filtroIdFNC" name="filtroIdFNC" placeholder="Número da FNC">
                </div>

                <div class="col-12 col-md-4 col-lg-3">
                    <label for="filtroPedido" class="form-label">Pedido</label>
                    <input type="text" class="form-control form-control-sm" id="filtroPedido" name="filtroPedido" placeholder="Número do Pedido">
                </div>

                <div class="col-12 col-md-4 col-lg-3">
                    <label for="filtroNotaFiscal" class="form-label">Nota Fiscal</label>
                    <input type="number" class="form-control form-control-sm" id="filtroNotaFiscal" name="filtroNotaFiscal" placeholder="Número da NF">
                </div>

                <div class="col-12 col-md-4 col-lg-3">
                    <label for="filtroSetor" class="form-label">Setor</label>
                    <select multiple class="form-select form-select-sm select2" id="filtroSetor" name="filtroSetor" data-placeholder="Todos os setores">
                        <option value="">Todos os setores</option>
                    </select>
                </div>

                <div class="col-12 col-md-4 col-lg-3">
                    <label for="filtroSituacao" class="form-label">Situação</label>
                    <select multiple class="form-select form-select-sm select2" id="filtroSituacao" name="filtroSituacao" data-placeholder="Todas">
                        <option value="">Todas</option>
                        <option value="1">ABERTA</option>
                        <option value="2">RESPONDIDA</option>
                        <option value="3">ENCERRADA</option>
                    </select>
                </div>

                <div class="col-12 col-md-4 col-lg-3">
                    <label for="filtroValidacao" class="form-label">Validação</label>
                    <select multiple class="form-select form-select-sm select2" id="filtroValidacao" name="filtroValidacao" data-placeholder="Todas">
                        <option value="">Todas</option>
                        <option value="1">REGISTRADA</option>
                        <option value="2">EM ANÁLISE</option>
                        <option value="3">PROCEDE</option>
                        <option value="4">NÃO PROCEDE</option>
                        <option value="5">CANCELADA</option>
                    </select>
                </div>

                <div class="col-12 col-md-4 col-lg-3">
                    <label for="filtroDisposicao" class="form-label">Disposição</label>
                    <select multiple class="form-select form-select-sm select2" id="filtroDisposicao" name="filtroDisposicao" data-placeholder="Todas">
                        <option value="">Todas</option>
                    </select>
                </div>

                <div class="col-12 col-md-4 col-lg-3">
                    <label for="filtroTipoDefeito" class="form-label">Área</label>
                    <select multiple class="form-select form-select-sm select2" id="filtroTipoDefeito" name="filtroTipoDefeito" data-placeholder="Todos">
                        <option value="">Todos</option>
                        <option value="ADM">ADMINISTRATIVO</option>
                        <option value="PRO">INDUSTRIAL</option>
                    </select>
                </div>

                <div class="col-12 col-md-4 col-lg-3">
                    <label for="filtroDefeito" class="form-label">Defeito</label>
                    <select multiple class="form-select form-select-sm select2" id="filtroDefeito" name="filtroDefeito" data-placeholder="Todos os defeitos">
                        <option value="">Todos os defeitos</option>
                    </select>
                </div>

                <div class="col-12 col-md-4 col-lg-3">
                    <label for="filtroResponsavel" class="form-label">Responsável Setorial</label>
                    <select multiple class="form-select form-select-sm select2" id="filtroResponsavel" name="filtroResponsavel" data-placeholder="Todos os responsáveis">
                        <option value="">Todos os responsáveis</option>
                    </select>
                </div>

                <div class="col-12 d-flex justify-content-end align-items-end gap-2">
                    <button type="button" class="btn btn-sm btn-outline-secondary" id="btnLimpar">
                        <i class="bi bi-x-lg me-1"></i> Limpar
                    </button>
                    <button type="button" class="btn btn-sm btn-primary" id="btnPesquisar">
                        <i class="bi bi-search me-1"></i> Pesquisar
                    </button>
                </div>
            </div>
        </div>

        <div id="avisoFilial" class="d-none alert alert-warning mb-4">
            Atenção! Você está visualizando dados de todas as filiais. Para ver dados de uma filial específica, selecione-a no filtro acima.
        </div>

        <div class="card">
           <div class="d-flex justify-content-between align-items-center p-3 border-bottom">
                <div class="d-flex gap-2">
                    <button type="button" class="btn btn-success btn-sm" id="btnExportarExcel">
                        <i class="bi bi-file-earmark-excel me-1"></i> Exportar Excel
                    </button>
                    <button type="button" class="btn btn-danger btn-sm" id="btnExportarPDF">
                        <i class="bi bi-file-earmark-pdf me-1"></i> Exportar PDF
                    </button>
                </div>
            </div>

            <div class="table-wrapper">
                <div id="loadingOverlay" class="loading-overlay" style="display: none;">
                    <div class="spinner-border text-primary" role="status">
                        <span class="visually-hidden">Carregando...</span>
                    </div>
                </div>

                <table class="table table-hover mb-0" id="gradeFNC">
                    <thead>
                        <tr>
                            <th>FNC</th>
                            <th>Filial</th>
                            <th>Data de Inclusão</th>
                            <th>Data de Produção</th>
                            <th>Data da Ocorrência</th>
                            <th>Data de Envio ao Cliente</th>
                            <th>Data de Conclusão Prevista</th>
                            <th>Qtde.</th>
                            <th>Situação</th>
                            <th>Status</th>
                            <th>Validação</th>
                            <th>Cód. Cliente</th>
                            <th>Cód. Loja</th>
                            <th>Razão Social</th>
                            <th>Cidade / Estado</th>
                            <th>Cód. Produto</th>
                            <th>Referência</th>
                            <th>Grupo Produto</th>
                            <th>Máquina</th>
                            <th>Turno</th>
                            <th>Setor</th>
                            <th>Defeito</th>
                            <th>Área</th>
                            <th>Disposição</th> 
                            <th>Descrição Disposição</th> 
                            <th>Eficácia</th> 
                            <th>Nota Fiscal SW</th>  
                            <th>Coeficiente NQA</th>
                            <th>Valor Agregado</th> 
                            <th>Valor Reclamado</th> 
                            <th>Preço Unitário</th>  
                            <th>Qtde. Vendida</th>         
                            <th>Total NF</th>      
                            <th>Data Faturamento</th>           
                            <th>Analista</th>  
                            <th>Responsável</th> 
                            <th>Dias em Aberto até Resposta</th>      
                            <th>Dias em Aberto até Encerramento</th>     
                            <th>Ações</th>                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                    
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

    <div class="modal fade" id="modalRespostas" tabindex="-1" aria-labelledby="modalRespostasLabel" aria-hidden="true">
        <div class="modal-dialog modal-xl">
            <div class="modal-content">
                <div class="modal-header">
                    <h5 class="modal-title" id="modalRespostasLabel">
                        <i class="fas fa-history me-2"></i>Respostas
                    </h5>
                    <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Close"></button>
                </div>
                <div class="modal-body" id="modalRespostasBody"></div>
                <div class="modal-footer">
                    <button type="button" class="btn btn-secondary btn-sm" data-bs-dismiss="modal">Fechar</button>
                </div>
            </div>
        </div>
    </div>

    <div class="modal fade" id="modalObservacoes" tabindex="-1" aria-labelledby="modalObservacoesLabel" aria-hidden="true">
        <div class="modal-dialog modal-xl">
            <div class="modal-content">
                <div class="modal-header">
                    <h5 class="modal-title" id="modalObservacoesLabel">
                        <i class="fas fa-history me-2"></i>Observações
                    </h5>
                    <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Close"></button>
                </div>
                <div class="modal-body" id="modalObservacoesBody"></div>
                <div class="modal-footer">
                    <button type="button" class="btn btn-secondary btn-sm" data-bs-dismiss="modal">Fechar</button>
                </div>
            </div>
        </div>
    </div>

</body>
</html>
