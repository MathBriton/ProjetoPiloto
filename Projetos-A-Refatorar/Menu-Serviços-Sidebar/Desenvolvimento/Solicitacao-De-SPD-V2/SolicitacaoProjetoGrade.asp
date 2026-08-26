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

usuarioEhRepresentante = 0
If SESSION("NIVEL") = "1" Then
    usuarioEhRepresentante = 1
End If

nivelSelecionado = Request("nivelSelecionado")
If nivelSelecionado <> "" Then
    SESSION("NIVELSPD") = nivelSelecionado
End If

nivelUsuario = SESSION("NIVELSPD")
%>

<!DOCTYPE html>
<html lang="pt-br">
<head>
    <title>Solicitações de Projeto ao Desenvolvimento (SPD)</title>
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

        #gradeSolicitacao th,
        #gradeSolicitacao td {
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
            background-color: #e9d5ff;
            color: #6b21a8;
        }
        .badge-blue {
            background-color: #cffafe;
            color: #155e75;
        }
        .badge-yellow {
            background-color: #fef3c7;
            color: #92400e;
        }
        .badge-orange {
            background-color: #fed7aa;
            color: #c2410c;
        }
        .badge-gray {
            background-color: #e2e2e4;
            color: #374151;
        }
        .truncate-text {
            max-width: 200px;
            white-space: nowrap;
            overflow: hidden;
            text-overflow: ellipsis;
            cursor: help;
        }
        .z-index-9999 {
            z-index: 9999 !important;
        }

        .barra-teste {
            background-color: #fef9c3;
            border-bottom: 2px solid #facc15;
            padding: 6px 16px;
            display: flex;
            align-items: center;
            gap: 10px;
            font-size: .8rem;
            font-weight: 600;
            color: #713f12;
            position: sticky;
            top: 0;
            z-index: 1050;
        }

        .barra-teste select {
            font-size: .8rem;
            font-weight: 600;
            padding: 2px 8px;
            border-radius: 6px;
            border: 1.5px solid #facc15;
            background-color: #fff;
            color: #713f12;
            cursor: pointer;
        }

        .barra-teste select:focus {
            outline: none;
            box-shadow: 0 0 0 2px #fde68a;
        }

        .barra-teste .nivel-badge {
            display: inline-flex;
            align-items: center;
            gap: 5px;
            background-color: #fde68a;
            border-radius: 20px;
            padding: 2px 10px;
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
    <script src="../../components/toast.js?v=<%=Now()%>"></script>
    <script src="../../components/confirmation.js?v=<%=Now()%>"></script>
    <script src="/corporativo/funcoes/funcoesJAVASCRIPT.js?v=<%=Now()%>"></script>

    <script>
        const eNivel = Object.freeze({
            COMERCIAL: 'C',
            DESENVOLVIMENTO: 'D',
            PCP: 'P'
        })

        const eStatus = Object.freeze({
            PENDENTE_COMERCIAL: 'S',
            BLOQUEADO_COMERCIAL: 'V',
            PENDENTE_DESENVOLVIMENTO: 'A',
            PENDENTE_PCP: 'D',
            //APROVADO_PCP: 'F',
            BLOQUEADO_PCP: 'G',
            EXPIRADO: 'E',
            APROVADO: 'P',
            REPROVADO: 'X',
            CANCELADO: 'K',
        });

        const eStatusAndamento = Object.freeze({
            AMOSTRA_ENVIADA: 'E',
            AMOSTRA_APROVADA: 'A',
            AMOSTRA_REPROVADA: 'R',
            ANALISE_DO_PROJETO: 'P',
            ENVIO_ESPECIFICACAO_AMOSTRA: 'G',
            LIBERACAO_PPCP_PRODUCAO: 'B',
            APROVACAO_LAYOUT: 'L',
            LIBERACAO_COMERCIAL_GERENCIA: 'C',
            LIBERACAO_COMERCIAL_DIRETORIA: 'D',
            CADASTRO_FT: 'F',
            PADRAO_DE_COR: 'H',
            CONFECCAO_FERRAMENTAL: 'J',
            FT_LIBERADA: 'K',
            SEM_RETORNO: 'T',
            LIBERACAO_ANALISE_CUSTO: 'I',
            APROVACAO_CUSTOS_GERENCIA: 'M',
            DESENVOLVIMENTO_LAYOUT: 'N',
            CORTE_AMOSTRAS: 'O',
            FALTA_CHAPAS: 'Q',
            ORCAMENTO: 'S',
            BPP_BETTER_PLANET_PACKING: 'U',
            ECOMMERCE_ESHOP: 'V'
        });

        const eProcesso = Object.freeze({
            INOVACAO: '1',
            MELHORIA_CONTINUA_SK: '4',
            MELHORIA_CONTINUA_CLIENTE: '5',
            ALTERACAO_LAYOUT: '6',
            ALTERACAO_DIMENSIONAL: '7',
            ALTERACAO_ESTRUTURA: '8',
            NOVA_REFERENCIA: '9'
        });

        const usuarioEhRepresentante = <%= usuarioEhRepresentante %>;
        let nivelUsuario = '<%= nivelUsuario %>';
        
        let dadosSolicitacoes = [];
        let carregando = false;

        function alterarNivelTeste(valor) {
            const codPlanta = $('#filtroFilial')?.val() || '';
            window.location.href = `SolicitacaoProjetoGrade.asp?nivelSelecionado=${valor}&codPlanta=${codPlanta}` 
        }

        function exibirLoading(exibir) {
            carregando = exibir;

            if (exibir) {
                $('#loadingOverlay').show();
            } else {
                $('#loadingOverlay').hide();
            }
        }

        function carregarSolicitacoes() {    
            if (!validarFiltros()) 
                return
            
            exibirLoading(true);

            const params = {
                acao: 'listar',
                filtroFilial: $('#filtroFilial').val(),
                filtroAndamentoProcesso: $('#filtroAndamentoProcesso').val(),
                filtroStatus: $('#filtroStatus').val(),
                filtroCodCliente: $('#filtroCodCliente').val(),
                filtroCodLoja: $('#filtroCodLoja').val(),
                filtroIdSolicitacao: $('#filtroIdSolicitacao').val(),
                filtroFtOrigem: $('#filtroFtOrigem').val(),
                filtroFtNova: $('#filtroFtNova').val(),
                flagApenasPendentes: $('#flagApenasPendentes').is(':checked') ? 1 : 0
            };

            $.ajax({
                url: 'SolicitacaoProjetoRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: params,
                dataType: 'json',
                success: function(response) {
                    exibirLoading(false);

                    if (!response.sucesso) {
                        Toast.error(response.mensagem);
                        return;
                    }
                
                    dadosSolicitacoes = response.dados || [];
                    renderizarGrade();
                },
                error: function(xhr, status, error) {
                    exibirLoading(false);
                    console.error('Erro ao pesquisar:', error);
                    Toast.error('Erro ao realizar a pesquisa');
                }
            });
        }

        function carregarAndamentosProcesso() {
            const $select = $('#filtroAndamentoProcesso');
            $select.empty().append('<option value="">Todos os andamentos</option>');

            const params = {
                acao: 'listarStatusAndamento'
            };

            $.ajax({
                url: 'SolicitacaoProjetoRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: params,
                dataType: 'json',
                success: function(response) {
                    if (!response.sucesso) {
                        Toast.error("Erro ao carregar a lista de andamentos do processo: " + response.mensagem);
                        return;
                    }

                    response.dados.forEach(function(item) {
                        $select.append($('<option>', { value: item.codigo, text: item.descricao }));
                    });
                },
                error: function(xhr, status, error) {
                    Toast.error('Erro ao carregar a lista de andamentos do processo:', error);
                }
            });
        }

        function renderizarGrade() {
            const $tbody = $('#gradeSolicitacao tbody');
            
            $tbody.empty();
            
            if (dadosSolicitacoes.length === 0) {
                $tbody.html(`
                    <tr>
                        <td colspan="28" class="text-center text-muted py-4">
                            Nenhum registro encontrado
                        </td>
                    </tr>
                `);
                return;
            }

            dadosSolicitacoes.forEach(function(item) {    
                const statusBadge = obterBadgeStatus(item.status);
                const processoBadge = obterBadgeProcesso(item.processo);                
                const andamentoBadge = obterBadgeStatusAndamento(item.statusAndamento);                
                
                const acoesMap = {
                   [eNivel.COMERCIAL]: obterAcoesComercial(item),
                   [eNivel.DESENVOLVIMENTO]: obterAcoesDesenvolvimento(item),
                   [eNivel.PCP]: obterAcoesPCP(item)
                }
                
                let htmlAcoes = ''
                if(usuarioEhRepresentante) {
                    htmlAcoes = obterAcoesRepresentante(item);
                } else {
                    htmlAcoes = acoesMap[nivelUsuario] || '';
                }
                
                $tbody.append(`
                    <tr>
                        <td>${item.id || '-'}</td>
                        <td>${item.codCliente} - ${item.nomeCliente}</td>
                        <td>${item.nomeVendedor || '-'}</td>
                        <td>
                            <div class="truncate-text" title="${item.referencia || ''}">
                                ${item.referencia || '-'}
                            </div>
                        </td>
                        <td>${item.ftNova || '-'}</td>
                        <td>${statusBadge}</td>
                        <td>${processoBadge}</td>
                        <td>${andamentoBadge}</td>
                        <td>${item.dataEmissao || '-'}</td>
                        ${htmlAcoes}
                    </tr>
                `);
            });
        }
   
        function novaSolicitacao() {
            window.location.href = `SolicitacaoProjetoFormulario.asp`;
        }

        function visualizarSolicitacao(id, codPlanta) {
            window.location.href = `SolicitacaoProjetoFormulario.asp?id=${id}&codPlanta=${codPlanta}&visualizar=1`;
        }

        function editarSolicitacao(id, codPlanta) {
            window.location.href = `SolicitacaoProjetoFormulario.asp?id=${id}&codPlanta=${codPlanta}`;
        }

        function alterarStatusAndamento(id, codPlanta) {
            window.location.href = `SolicitacaoProjetoAlterarStatusAndamento.asp?id=${id}&codPlanta=${codPlanta}`;
        }

        function editarTestesFisicos(id, codPlanta) {
            window.location.href = `SolicitacaoProjetoAlterarTestesFisicos.asp?id=${id}&codPlanta=${codPlanta}`;
        }

        function abrirEspelhoAmostra(id, codPlanta) {
            window.open(`../EspelhoPedidoAmostra.asp?Id=${id}&Planta=${codPlanta}`, '_blank');
        }
        
        function abrirModalHistoricoAlteracoes(id, codPlanta) {
            $('#idHistorico').text(id);
            const modal = new bootstrap.Modal(document.getElementById('modalHistoricoAlteracoes'));
            modal.show();

            const $tbody = $('#gradeHistorico tbody');
            $tbody.empty();

            const listaAlteracoes = dadosSolicitacoes.find(item => item.id === id).alteracoes;
            
            if (listaAlteracoes.length === 0) {
                $tbody.html(`
                    <tr>
                        <td colspan="28" class="text-center text-muted py-4">
                            Nenhum registro encontrado
                        </td>
                    </tr>
                `);
                return;
            }

            listaAlteracoes.forEach(function(item) {    
                $tbody.append(`
                    <tr>
                        <td>${item.IdHistorico || '-'}</td>
                        <td>${obterStatusAndamentoTraduzido(item.IdProcesso) || '-'}</td>
                        <td>${item.DataInsercaoProcesso ? formatarParaDataBR(item.DataInsercaoProcesso) : '-'}</td>
                        <td>${item.DataAlteracaoProcesso ? formatarParaDataBR(item.DataAlteracaoProcesso) : '-'}</td>
                        <td>${item.LeadTime}</td>
                        <td>${item.QuemAlterou || '-'}</td>
                        <td>
                            <div class="truncate-text" title="${item.Complemento}">
                                ${item.Complemento}
                            </div>
                        </td>
                    </tr>
                `);
            });
        }

        function excluirSolicitacao(id, codPlanta) {
            ConfirmationModal.danger({
                message: 'Deseja realmente excluir essa solicitação?',
                confirmText: 'Confirmar',
                cancelText: 'Cancelar',
                onConfirm: () => {
                    confirmarExclusao(id, codPlanta);
                },
                onCancel: () => {}
            });
        }

        function cancelarSolicitacao(id, codPlanta) {
            ConfirmationModal.danger({
                message: 'Deseja realmente cancelar solicitação?',
                confirmText: 'Confirmar',
                cancelText: 'Cancelar',
                onConfirm: () => {
                    confirmarCancelamento(id, codPlanta);
                },
                onCancel: () => {}
            });
        }

        function reverterCancelamento(id, codPlanta) {
            ConfirmationModal.danger({
                message: 'Deseja realmente reverter o cancelamento dessa solicitação?',
                confirmText: 'Confirmar',
                cancelText: 'Cancelar',
                onConfirm: () => {
                    confirmarReversao(id, codPlanta);
                },
                onCancel: () => {}
            });
        }

        function avaliarSolicitacao(id, codPlanta) {
            window.location.href = `SolicitacaoProjetoFormularioAvaliacao.asp?id=${id}&codPlanta=${codPlanta}`;
        }

        function confirmarExclusao(id, codPlanta) {            
            exibirLoading(true);

            const params = {
                acao: 'excluir',
                id: id,
                codPlanta: codPlanta
            };

            $.ajax({
                url: 'SolicitacaoProjetoRepositorio.asp?v=<%=Now()%>',
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
                    carregarSolicitacoes();
                },
                error: function(xhr, status, error) {
                    exibirLoading(false);
                    console.error('Erro ao excluir:', error);
                    Toast.error('Erro ao excluir');
                }
            });
        }

        function confirmarReversao(id, codPlanta) {            
            exibirLoading(true);

            const params = {
                acao: 'reverterCancelamento',
                id: id,
                codPlanta: codPlanta
            };

            $.ajax({
                url: 'SolicitacaoProjetoRepositorio.asp?v=<%=Now()%>',
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
                    carregarSolicitacoes();
                },
                error: function(xhr, status, error) {
                    exibirLoading(false);
                    console.error('Erro ao reverter:', error);
                    Toast.error('Erro ao reverter');
                }
            });
        }

        function confirmarCancelamento(id, codPlanta) {            
            exibirLoading(true);

            const params = {
                acao: 'cancelar',
                id: id,
                codPlanta: codPlanta
            };

            $.ajax({
                url: 'SolicitacaoProjetoRepositorio.asp?v=<%=Now()%>',
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
                    carregarSolicitacoes();
                },
                error: function(xhr, status, error) {
                    exibirLoading(false);
                    console.error('Erro ao cancelar:', error);
                    Toast.error('Erro ao cancelar');
                }
            });
        }

        function exportarExcel() {
            if (dadosSolicitacoes.length === 0) {
                alert('Não há dados para exportar');
                return;
            }

            const nomeArquivo = 'Relatorio_FNC.csv';
            const colunas = [
            ];

            let csvContent = '\uFEFF'; 
            csvContent += colunas.join(';') + '\r\n';

            dadosSolicitacoes.forEach(item => {
                const linha = [
                    item.idFNC || '-',
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
        
        function pesquisarCliente() {
            const codPlanta = $('#filtroFilial').val();
            const baseUrl = `${location.protocol}//${location.host}`;
            const url = `${baseUrl}/corporativo/buscaCLIENTE.asp?id1=filtroCodCliente&id2=filtroCodLoja&id3=nomeCliente&id6=${codPlanta}`;
            window.open(url, '', "width=650,height=400");
        }
  
        function validarFiltros() {
            if (!$('#filtroFilial').val()) {
                $('#filtroFilial').focus();
                return false;
            }
            return true;
        }

        function limparFiltros() {
            $('#filtroFilial').val('').trigger('change');
            $('#filtroAndamentoProcesso').val('').trigger('change');
            $('#filtroStatus').val('').trigger('change');
            $('#filtroCodCliente').val('');
            $('#filtroCodLoja').val('');
            $('#filtroIdSolicitacao').val('');
            $('#filtroFtOrigem').val('');
            $('#filtroFtNova').val('');

            dadosSolicitacoes = [];
            renderizarGrade();
        }

        function obterBadgeStatus(status) {
            switch(status) {
                case eStatus.AGUARDANDO_COMERCIAL:
                    return '<span class="badge badge-blue">Aguardando Comercial</span>';
                case eStatus.AGUARDANDO_PCP:
                    return '<span class="badge badge-blue">Aguardando PCP</span>';
                case eStatus.APROVADO_PCP:
                    return '<span class="badge badge-orange">Aprovado PCP</span>';
                case eStatus.BLOQUEADO_PCP:
                    return '<span class="badge badge-red">Bloqueado PCP</span>';
                case eStatus.EXPIRADO:
                    return '<span class="badge badge-purple">Expirado</span>';
                case eStatus.APROVADO:
                    return '<span class="badge badge-green">Aprovado Total</span>';
                case eStatus.REPROVADO:
                    return '<span class="badge badge-red">Reprovado</span>';
                case eStatus.CANCELADO:
                    return '<span class="badge badge-gray">Cancelado</span>';
                case eStatus.BLOQUEADO_COMERCIAL:
                    return '<span class="badge badge-red">Bloqueado Comercial</span>';
                case eStatus.APROVADO_COMERCIAL:
                    return '<span class="badge badge-yellow">Aprovado Comercial</span>';
                default:
                    return '-';
            }
        }

        function obterAcoesRepresentante(item) {
            const itens = [
                `<li><a class="dropdown-item" onclick="visualizarSolicitacao('${item.id}', '${item.codPlanta}')"><i class="bi bi-eye me-1"></i> Visualizar</a></li>`,
            ];

            return `
                <td style="text-align:center;">
                    <button class="btn btn-sm btn-primary dropdown-toggle" type="button" data-bs-toggle="dropdown">
                        <i class="bi bi-three-dots-vertical"></i> Ações
                    </button>
                    <ul class="dropdown-menu z-3">
                        ${itens.join('')}
                    </ul>
                </td>
            `;
        }

        function obterAcoesComercial(item) {
            const itens = [
                `<li><a class="dropdown-item" onclick="visualizarSolicitacao('${item.id}', '${item.codPlanta}')"><i class="bi bi-eye me-1"></i> Visualizar</a></li>`,
                `<li><hr class="dropdown-divider"></li>`,
            ];

            if(item.status == eStatus.AGUARDANDO_COMERCIAL) {
                itens.push(`<li><a class="dropdown-item" onclick="avaliarSolicitacao('${item.id}', '${item.codPlanta}')"><i class="bi bi-clipboard-data me-1"></i> Aprovar ou Rejeitar</a></li>`)
            }

            return `
                <td style="text-align:center;">
                    <button class="btn btn-sm btn-primary dropdown-toggle" type="button" data-bs-toggle="dropdown">
                        <i class="bi bi-three-dots-vertical"></i> Ações
                    </button>
                    <ul class="dropdown-menu z-3">
                        ${itens.join('')}
                    </ul>
                </td>
            `;
        }

        function obterAcoesPCP(item) {
            const itens = [
                `<li><a class="dropdown-item" onclick="visualizarSolicitacao('${item.id}', '${item.codPlanta}')"><i class="bi bi-eye me-1"></i> Visualizar</a></li>`,
                `<li><hr class="dropdown-divider"></li>`,
            ];

            if(item.status == eStatus.AGUARDANDO_PCP) {
                itens.push(`<li><a class="dropdown-item" onclick="avaliarSolicitacao('${item.id}', '${item.codPlanta}')"><i class="bi bi-clipboard-data me-1"></i> Aprovar ou Rejeitar</a></li>`)
            }

            return `
                <td style="text-align:center;">
                    <button class="btn btn-sm btn-primary dropdown-toggle" type="button" data-bs-toggle="dropdown">
                        <i class="bi bi-three-dots-vertical"></i> Ações
                    </button>
                    <ul class="dropdown-menu z-3">
                        ${itens.join('')}
                    </ul>
                </td>
            `;
        }

        function obterAcoesDesenvolvimento(item) {
            const itens = [
                `<li><a class="dropdown-item" onclick="visualizarSolicitacao('${item.id}', '${item.codPlanta}')"><i class="bi bi-eye me-1"></i> Visualizar</a></li>`,
                `<li><a class="dropdown-item" onclick="abrirModalHistoricoAlteracoes('${item.id}')"><i class="bi bi-clock me-1"></i> Histórico de Alterações</a></li>`,
            ];

            if (item.quantidadeAmostras > 0 ) {                
                itens.push( `<li><a class="dropdown-item" onclick="abrirEspelhoAmostra('${item.id}', '${item.codPlanta}')"><i class="bi bi-image me-1"></i> Espelho de Amostra</a></li>`)
            }

            if (item.status == eStatus.APROVADO_PCP) {
                itens.push(`<li><a class="dropdown-item" onclick="avaliarSolicitacao('${item.id}', '${item.codPlanta}')"><i class="bi bi-clipboard-data me-1"></i> Aprovar ou Rejeitar</a></li>`)
            }
           
            if([eStatus.APROVADO_COMERCIAL, eStatus.APROVADO_PCP, eStatus.BLOQUEADO_PCP].includes(item.status)) {
                itens.push(`<li><a class="dropdown-item" onclick="editarSolicitacao('${item.id}', '${item.codPlanta}')"><i class="bi bi-pencil-square me-1"></i> Editar</a></li>`);
                itens.push(`<li><a class="dropdown-item" onclick="cancelarSolicitacao('${item.id}', '${item.codPlanta}')"><i class="bi bi-ban"></i> Cancelar Solicitação</a></li>`);
                itens.push(`<li><a class="dropdown-item" onclick="editarTestesFisicos('${item.id}', '${item.codPlanta}')"><i class="bi bi-box-seam me-1"></i> Editar Testes</a></li>`);
                itens.push(`<li><a class="dropdown-item" onclick="alterarStatusAndamento('${item.id}', '${item.codPlanta}')"><i class="bi bi-clipboard-check me-1"></i> Alterar Andamento</a></li>`);
                itens.push(`<li><hr class="dropdown-divider"></li>`);
                itens.push(`<li><a class="dropdown-item text-danger" onclick="excluirSolicitacao('${item.id}', '${item.codPlanta}')"><i class="bi bi-trash me-1"></i> Excluir</a></li>`);
            }

            if(item.status == eStatus.CANCELADO) {
                itens.push(`<li><a class="dropdown-item" onclick="reverterCancelamento('${item.id}', '${item.codPlanta}')"><i class="bi bi-arrow-counterclockwise"></i> Reverter Cancelamento</a></li>`);
            }

            return `
                <td style="text-align:center;">
                    <button class="btn btn-sm btn-primary dropdown-toggle" type="button" data-bs-toggle="dropdown">
                        <i class="bi bi-three-dots-vertical"></i> Ações
                    </button>
                    <ul class="dropdown-menu z-3">
                        ${itens.join('')}
                    </ul>
                </td>
            `;
        }

        function obterBadgeStatusAndamento(statusAndamento) {
            switch(statusAndamento) {
                case eStatusAndamento.AMOSTRA_ENVIADA:
                    return '<span class="badge badge-green">Amostra Enviada</span>';
                case eStatusAndamento.AMOSTRA_APROVADA:
                    return '<span class="badge badge-green">Amostra Aprovada</span>';
                case eStatusAndamento.REPROVADA:
                    return '<span class="badge badge-red">Amostra Reprovada</span>';
                case eStatusAndamento.ANALISE_PROJETO:
                    return '<span class="badge badge-yellow">Análise do Projeto</span>';
                case eStatusAndamento.ESPECIFICACAO_AMOSTRA:
                    return '<span class="badge badge-yellow">Envio de Especif./Amostra</span>';
                case eStatusAndamento.ENVIO_ESPECIFICACAO_AMOSTRA:
                    return '<span class="badge badge-yellow">Lib. PPCP/Produção</span>';
                case eStatusAndamento.APROVACAO_LAYOUT:
                    return '<span class="badge badge-yellow">Aprovação de Layout</span>';
                case eStatusAndamento.LIBERACAO_COMERCIAL_GERENCIA:
                    return '<span class="badge badge-yellow">Lib. Comercial - Gerência</span>';
                case eStatusAndamento.LIBERACAO_COMERCIAL_DIRETORIA:
                    return '<span class="badge badge-yellow">Lib. Comercial - Diretoria</span>';
                case eStatusAndamento.CADASTRO_FT:
                    return '<span class="badge badge-yellow">Cadastro de FT</span>';
                case eStatusAndamento.PADRAO_DE_COR:
                    return '<span class="badge badge-yellow">Padrão de Cor</span>';
                case eStatusAndamento.CONFECCAO_FERRAMENTAL:
                    return '<span class="badge badge-yellow">Confecção do Ferramental</span>';
                case eStatusAndamento.FT_LIBERADA:
                    return '<span class="badge badge-yellow">FT Liberada</span>';
                case eStatusAndamento.SEM_RETORNO:
                    return '<span class="badge badge-yellow">Sem Retorno</span>';
                case eStatusAndamento.LIBERACAO_ANALISE_CUSTO:
                    return '<span class="badge badge-yellow">Lib. Análise de Custo</span>';
                case eStatusAndamento.APROVACAO_CUSTOS_GERENCIA:
                    return '<span class="badge badge-yellow">Aprovação Custos Gerência</span>';
                case eStatusAndamento.DESENVOLVIMENTO_LAYOUT:
                    return '<span class="badge badge-yellow">Desenvolvimento de Layout</span>';
                case eStatusAndamento.CORTE_AMOSTRAS:
                    return '<span class="badge badge-yellow">Corte de Amostras</span>';
                case eStatusAndamento.FALTA_CHAPAS:
                    return '<span class="badge badge-yellow">Falta de Chapas</span>';
                case eStatusAndamento.ORCAMENTO:
                    return '<span class="badge badge-yellow">Orçamento</span>';
                case eStatusAndamento.BPP_BETTER_PLANET_PACKING:
                    return '<span class="badge badge-yellow">BPP - Better Planet Packing</span>';
                case eStatusAndamento.ECOMMERCE_ESHOP:
                    return '<span class="badge badge-yellow">E-commerce/Eshop</span>';
                default:
                    return '-';
            }
        }

        function obterBadgeProcesso(processo) {
            switch(processo) {
                case eProcesso.INOVACAO:
                    return '<span class="badge badge-green">Inovação</span>';
                case eProcesso.MELHORIA_CONTINUA_SK: 
                    return '<span class="badge badge-green">Melhoria Continua - SK</span>';
                case eProcesso.MELHORIA_CONTINUA_CLIENTE:
                    return '<span class="badge badge-red">Melhoria Continua - Cliente</span>';
                case eProcesso.ALTERACAO_LAYOUT: 
                    return '<span class="badge badge-yellow">Alteração Layout</span>';
                case eProcesso.ALTERACAO_DIMENSIONAL: 
                    return '<span class="badge badge-yellow">Alteração Dimensional</span>';
                case eProcesso.ALTERACAO_ESTRUTURA: 
                    return '<span class="badge badge-yellow">Alteração Estrutura</span>';
                case eProcesso.NOVA_REFERENCIA: 
                    return '<span class="badge badge-yellow">Nova Referência conforme especificado</span>';
                default:
                    return '-';
            }
        }

        function obterStatusAndamentoTraduzido(statusAndamento) {
            switch (statusAndamento) {
                case eStatusAndamento.AMOSTRA_ENVIADA:
                    return "Amostra Enviada" 
                case eStatusAndamento.AMOSTRA_APROVADA: 
                    return "Amostra Aprovada"
                case eStatusAndamento.AMOSTRA_REPROVADA: 
                    return "Amostra Reprovada"
                case eStatusAndamento.ANALISE_DO_PROJETO: 
                    return "Análise do Projeto"
                case eStatusAndamento.ENVIO_ESPECIFICACAO_AMOSTRA: 
                    return "Envio de Especif./Amostra"
                case eStatusAndamento.LIBERACAO_PPCP_PRODUCAO: 
                    return "Lib. PPCP/Produção"
                case eStatusAndamento.APROVACAO_LAYOUT: 
                    return "Aprovação de Layout"
                case eStatusAndamento.LIBERACAO_COMERCIAL_GERENCIA: 
                    return "Lib. Comercial - Gerência"
                case eStatusAndamento.LIBERACAO_COMERCIAL_DIRETORIA: 
                    return "Lib. Comercial - Diretoria"
                case eStatusAndamento.CADASTRO_FT: 
                    return "Cadastro de FT"
                case eStatusAndamento.PADRAO_DE_COR: 
                    return "Padrão de Cor"
                case eStatusAndamento.CONFECCAO_FERRAMENTAL: 
                    return "Confecção do Ferramental"
                case eStatusAndamento.FT_LIBERADA: 
                    return "FT Liberada"
                case eStatusAndamento.SEM_RETORNO: 
                    return "Sem Retorno"
                case eStatusAndamento.LIBERACAO_ANALISE_CUSTO: 
                    return "Lib. Análise de Custo"
                case eStatusAndamento.APROVACAO_CUSTOS_GERENCIA: 
                    return "Aprovação Custos Gerência"
                case eStatusAndamento.DESENVOLVIMENTO_LAYOUT: 
                    return "Desenvolvimento de Layout"
                case eStatusAndamento.CORTE_AMOSTRAS: 
                    return "Corte de Amostras"
                case eStatusAndamento.FALTA_CHAPAS: 
                    return "Falta de Chapas"
                case eStatusAndamento.ORCAMENTO: 
                    return "Orçamento"
                case eStatusAndamento.BPP_BETTER_PLANET_PACKING: 
                    return "BPP - Better Planet Packing"
                case eStatusAndamento.ECOMMERCE_ESHOP: 
                    return "E-commerce/Eshop"
                default:
                    return "Não Informado"
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

        $(document).ready(function() {
            inicializarSelect2();
            carregarAndamentosProcesso();
            renderizarGrade();

            const urlParams = new URLSearchParams(window.location.search);
            const codigoPlanta = urlParams.get('codPlanta') || '<%=SESSION("FILIAL")%>';

            if (urlParams.has('id')) 
                $('#filtroIdSolicitacao').val(urlParams.get('id'));

            if (usuarioEhRepresentante) {
                $('#divNovoRegistro').show();
            }

            $('#filtroFilial').val(codigoPlanta).trigger('change');
            carregarSolicitacoes();
        });
    </script>
</head>
<body>
    <div class="barra-teste">
        <i class="bi bi-wrench-adjustable"></i>
        <span>Modo de teste — Visualizando como:</span>
        <span class="nivel-badge">
            <i class="bi bi-person-fill"></i>
            <span id="labelNivelAtivo">
                <% 
                    If nivelUsuario = "D" Then
                        Response.Write "Desenvolvimento"
                    ElseIf nivelUsuario = "C" Then
                        Response.Write "Comercial"
                    ElseIf nivelUsuario = "P" Then
                        Response.Write "PCP"
                    End If
                %>
            </span>
        </span>
        <select onchange="alterarNivelTeste(this.value)">
            <option value="C" <%= Iif(nivelUsuario = "C", "selected", "") %>>Comercial</option>
            <option value="D" <%= Iif(nivelUsuario = "D", "selected", "") %>>Desenvolvimento</option>
            <option value="P" <%= Iif(nivelUsuario = "P", "selected", "") %>>PCP</option>
        </select>
    </div>

    <div class="container my-4">
        <div class="my-4">
            <h3>Solicitações de Projeto ao Desenvolvimento (SPD)</h3>
        </div>

        <div class="card p-3 mb-4">
            <div class="row g-3">
                <div class="col-12 col-md-4">
                    <label for="filtroFilial" class="form-label">Filial</label>
                    <select required class="form-select form-select-sm select2" id="filtroFilial" name="filtroFilial" data-placeholder="Todas as filiais">
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
                
                <div class="col-12 col-md-4">
                    <label for="filtroAndamentoProcesso" class="form-label">Andamento Processo</label>
                    <select required class="form-select form-select-sm select2" id="filtroAndamentoProcesso" name="filtroAndamentoProcesso" data-placeholder="Todos os andamentos">
                        <option value="">Todas as filiais</option>
                    </select>
                </div>

                <div class="col-12 col-md-4">
                    <label for="filtroCodCliente" class="form-label">Cód. Cliente</label>
                    <div class="input-group input-group-sm">
                        <button type="button" class="btn btn-sm btn-secondary" onclick="pesquisarCliente()">
                            Buscar <i class="bi bi-box-arrow-up-right"></i>
                        </button>
                        <input type="text" class="form-control form-control-sm" id="filtroCodCliente" name="filtroCodCliente" placeholder="Código do cliente">
                    </div>
                    <small class="text-muted">Clique no botão para buscar ou digite o código</small>
                </div>

                <div class="col-12 col-md-4">
                    <label for="filtroCodLoja" class="form-label">Cód. Loja</label>
                    <input type="text" class="form-control form-control-sm" id="filtroCodLoja" name="filtroCodLoja" placeholder="Código da loja">
                </div>

                <div class="col-12 col-md-4">
                    <label for="filtroStatus" class="form-label">Status</label>
                    <select required class="form-select form-select-sm select2" id="filtroStatus" name="filtroStatus" data-placeholder="Todos">
                        <option value="">Todos</option>
                        <option value="S">Aguardando Comercial</option>
                        <option value="A">Aprovado Comercial</option>
                        <option value="D">Aguardando PCP</option>
                        <option value="F">Aprovado PCP</option>
                        <option value="E">Expirado</option>
                        <option value="P">Aprovado Total</option>
                        <option value="X">Reprovado</option>
                        <option value="K">Cancelado</option>
                    </select>
                </div>

                <div class="col-12 col-md-4">
                    <label for="filtroIdSolicitacao" class="form-label">Id Solicitação</label>
                    <input type="number" class="form-control form-control-sm" id="filtroIdSolicitacao" name="filtroIdSolicitacao" placeholder="Id da solicitação">
                </div>

                <div class="col-12 col-md-4">
                    <label for="filtroFtOrigem" class="form-label">FT Copia</label>
                    <input type="number" class="form-control form-control-sm" id="filtroFtOrigem" name="filtroFtOrigem" placeholder="Número da FT de origem">
                </div>

                <div class="col-12 col-md-4">
                    <label for="filtroFtNova" class="form-label">FT</label>
                    <input type="number" class="form-control form-control-sm" id="filtroFtNova" name="filtroFtNova" placeholder="Número da FT">
                </div>

                <div class="col-12 mt-4 d-flex justify-content-between">
                    <div class="d-flex align-items-end">
                        <div class="form-check form-switch">
                            <input onchange="carregarSolicitacoes()" checked class="form-check-input" type="checkbox" name="flagApenasPendentes" id="flagApenasPendentes" />
                            <label class="form-check-label" for="flagApenasPendentes">Apenas minhas pendências</label>
                        </div>
                    </div>

                    <div class="d-flex justify-content-end align-items-end gap-2">
                        <button onclick="limparFiltros()" type="button" class="btn btn-sm btn-outline-secondary" id="btnLimpar">
                            <i class="bi bi-x-lg me-1"></i> Limpar
                        </button>
                        <button onclick="carregarSolicitacoes()" type="button" class="btn btn-sm btn-primary" id="btnPesquisar">
                            <i class="bi bi-search me-1"></i> Pesquisar
                        </button>
                    </div>
                </div>
            </div>
        </div>

        <div class="card">
            <div id="divNovoRegistro" class="gap-2 p-3 border-bottom">
                <button class="btn btn-sm btn-primary" onclick="novaSolicitacao()">
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

                <table class="table table-hover mb-0" id="gradeSolicitacao">
                    <thead>
                        <tr>
                            <th>#</th>
                            <th>Cliente</th>
                            <th>Vendedor</th>
                            <th>Referência</th>
                            <th>FT</th>
                            <th>Status</th>
                            <th>Processo</th>
                            <th>Andamento Processo</th>
                            <th>Data Emissão</th>    
                            <th></th>                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                    
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

    <!-- Modal de Histórico de atualizações -->
    <div class="modal fade" id="modalHistoricoAlteracoes" tabindex="-1" aria-labelledby="modalHistoricoAlteracoesLabel" aria-hidden="true">
        <div class="modal-dialog modal-xl">
            <div class="modal-content">
                <div class="modal-header">
                    <h5 class="modal-title" id="modalHistoricoAlteracoesTitulo">
                        Histórico de atualizações da Solicitação #<span id="idHistorico"></span>
                    </h5>
                    <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Fechar"></button>
                </div>
                <div class="modal-body">
                    <table class="table table-hover mb-0" id="gradeHistorico">
                        <thead>
                            <tr>
                                <th>#</th>
                                <th>Processo</th>
                                <th>Início</th>
                                <th>Fim</th>
                                <th>Leadtime</th>
                                <th>Quem</th>
                                <th>Complemento</th>
                            </tr>
                        </thead>
                        <tbody>
                            <tr>
                                <td class="text-center text-muted py-4">
                                    Nenhum registro encontrado
                                </td>
                            </tr>
                        </tbody>
                    </table>
                </div>
            </div>
        </div>
    </div>
</body>
</html>