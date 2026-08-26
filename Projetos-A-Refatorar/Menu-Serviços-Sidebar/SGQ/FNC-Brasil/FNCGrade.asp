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
    <title>Lista de FNC</title>
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

        #gradeFNC th,
        #gradeFNC td {
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
            REGISTRADA: '1',
            EM_ANALISE: '2',
            PROCEDE: '3',
            NAO_PROCEDE: '4',
            CANCELADA: '5'
        });

        const eSituacao = Object.freeze({
            ABERTA: '1',
            RESPONDIDA: '2',
            ENCERRADA: '3'
        })

        let nivelUsuario = '';
        let dadosFNC = [];
        let limitesNQA = {};

        function exibirLoading(exibir) {
            if (exibir) {
                $('#loadingOverlay').show();
            } else {
                $('#loadingOverlay').hide();
            }
        }

        function obterNivelUsuario() {
            return new Promise((resolve, reject) => {
                $.ajax({
                    url: 'FNCRepositorio.asp?v=<%=Now()%>',
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

        function carregarFNCs() { 
            exibirLoading(true);

            const params = {
                acao: 'listar',
                filtroFilial: $('#filtroFilial').val(),
                filtroDataInclusaoDe: $('#filtroDataInclusaoDe').val(),
                filtroDataInclusaoAte: $('#filtroDataInclusaoAte').val(),
                filtroStatus: $('#filtroStatus').val(),
                filtroSituacao: $('#filtroSituacao').val(),
                filtroAprovada: $('#filtroAprovada').val(),
                filtroCodCliente: $('#filtroCodCliente').val(),
                filtroCodLoja: $('#filtroCodLoja').val(),
                filtroIdFNC: $('#filtroIdFNC').val(),
                filtroTipoReclamacao: $('#filtroTipoReclamacao').val(),
                flagApenasPendentes: $('#flagApenasPendentes').is(':checked') ? 1 : 0
            };

            $.ajax({
                url: 'FNCRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: params,
                dataType: 'json',
                success: function(response) {
                    exibirLoading(false);

                    if (!response.sucesso) {
                        Toast.error(response.mensagem);
                        return;
                    }
                
                    dadosFNC = response.dados || [];
                    renderizarGrade();
                },
                error: function(xhr, status, error) {
                    exibirLoading(false);
                    console.error('Erro ao pesquisar:', error);
                    Toast.error('Erro ao realizar a pesquisa');
                }
            });
        }

        function carregarResponsaveis(codPlanta, responsavelId) {
            const $select = $('#responsavelIdEnvio');
            $select.empty().append('<option value="">Selecione um responsável</option>');

            const params = {
                acao: 'listarResponsaveis',
                codPlanta: codPlanta
            };

            $.ajax({
                url: 'FNCRepositorio.asp?v=<%=Now()%>',
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

                    $select.val(responsavelId).trigger('change');
                },
                error: function(xhr, status, error) {
                    console.error('Erro ao carregar responsáveis:', error);
                }
            });
        }

        function renderizarGrade() {
            const $tbody = $('#gradeFNC tbody');
            
            $tbody.empty();
            
            if (dadosFNC.length === 0) {
                $tbody.html(`
                    <tr>
                        <td colspan="28" class="text-center text-muted py-4">
                            Nenhum registro encontrado
                        </td>
                    </tr>
                `);
                return;
            }

            dadosFNC.forEach(function(item) {    
                const statusBadge = obterBadgeStatus(item.status);
                const situacaoBadge = obterBadgeSituacao(item.situacao);

                const acoesMap = {
                    'SOLICITANTE': obterAcoesSolicitante(item),
                    'ANALISTA': obterAcoesAnalista(item),
                    'RESPONSAVEL': obterAcoesResponsavel(item),
                    'APROVADOR': obterAcoesAprovador(item),
                }

                let htmlAcoes = acoesMap[nivelUsuario] || '';

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
                        <td>${item.id || '-'}</td>
                        <td>${item.dataInclusao || '-'}</td>
                        <td>${item.nomeFilial || '-'}</td>
                        <td>
                            ${item.tipoReclamacao  == 'NOTIFICACAO' ? 
                                `<span class="badge bg-warning text-dark">
                                    <i class="bi bi-exclamation-triangle-fill me-1"></i> Notificação
                                </span>` : 
                                `<span class="badge bg-success"> FNC </span>`
                            }
                        </td>
                        <td>${item.nomeSetor || '-'}</td>
                        <td>${item.nomeResponsavel || '-'}</td>
                        <td>${item.codCliente} - ${item.nomeCliente}</td></td>
                        <td>${item.codLoja}</td></td>
                        <td>${item.quantidadeReclamada || '-'}</td>
                        <td>${item.descricaoDisposicao || '-'}</td>
                        <td>${item.descricaoTipoDefeito || '-'}</td>
                        <td>${item.aprovada == '1' ? 'Sim' : 'Não'}</td>
                        <td>${statusNQA}</td>
                        <td>${statusBadge}</td>
                        <td>${situacaoBadge}</td>
                        ${htmlAcoes}
                    </tr>
                `);
            });
        }

        function novaFNC() {
            window.location.href = `FNCFormularioSolicitante.asp`;
        }

        function visualizarFNCSolicitante(id) {
            window.location.href = `FNCFormularioSolicitante.asp?id=${id}&acao=visualizar`;
        }

        function editarFNCSolicitante(id) {
            window.location.href = `FNCFormularioSolicitante.asp?id=${id}&acao=editar`;
        }

        function editarFNCAnalista(id) {
            window.location.href = `FNCFormularioAnalista.asp?id=${id}`;
        }
        
        function avaliarFNC(id) {
            window.location.href = `FNCFormularioAvaliacao.asp?id=${id}&acao=AVALIAR`;
        }

        function responderFNC(id) {
            window.location.href = `FNCFormularioAvaliacao.asp?id=${id}&acao=RESPONDER`;
        }

        function pesquisarCliente() {
            const codPlanta = $('#filtroFilial').val();
            if(!codPlanta) {
                Toast.warning("Para pesquisar por cliente é necessário selecionar uma filial")
                return;
            }

            const baseUrl = `${location.protocol}//${location.host}`;
            const url = `${baseUrl}/corporativo/buscaCLIENTE.asp?id1=filtroCodCliente&id2=filtroCodLoja&id3=hidNomeCliente&id6=${codPlanta}`;
            window.open(url, '', "width=650,height=400");
        }

        function limparFiltros() {
            $('#filtroFilial').val('').trigger('change');
            $('#filtroDataInclusaoDe').val('');
            $('#filtroDataInclusaoAte').val('');
            $('#filtroStatus').val('').trigger('change');
            $('#filtroSituacao').val('').trigger('change');
            $('#filtroAprovada').val('').trigger('change');
            $('#filtroCodCliente').val('');
            $('#filtroCodLoja').val('');
            $('#filtroIdFNC').val('');

            dadosFNC = [];
            renderizarGrade();
        }

        function obterBadgeStatus(status) {
            switch(status) {
                case eStatus.REGISTRADA:
                    return '<span class="badge badge-purple">REGISTRADA</span>';
                case eStatus.EM_ANALISE:
                    return '<span class="badge badge-yellow">EM ANÁLISE</span>';
                case eStatus.PROCEDE:
                    return '<span class="badge badge-green">PROCEDE</span>';
                case eStatus.NAO_PROCEDE:
                    return '<span class="badge badge-red">NÃO PROCEDE</span>';
                case eStatus.CANCELADA:
                    return '<span class="badge badge-gray">CANCELADA</span>';
                default:
                    return '-';
            }
        }

        function obterBadgeSituacao(status) {
            switch(status) {
                case eSituacao.ABERTA:
                    return '<span class="badge badge-green">ABERTA</span>';
                case eSituacao.RESPONDIDA:
                    return '<span class="badge badge-yellow">RESPONDIDA</span>';
                case eSituacao.ENCERRADA:
                    return '<span class="badge badge-pink">ENCERRADA</span>';
                default:
                    return '-';
            }
        }

        function obterAcoesSolicitante(item) {
            const itens = [
                `<a class="dropdown-item" onclick="visualizarFNCSolicitante('${item.id}')"><i class="bi bi-eye me-1"></i> Visualizar</a>`,
            ];

            if(item.tipoReclamacao == 'NOTIFICACAO') {
                if(item.status != eStatus.CANCELADA) {
                    itens.push(`<a class="dropdown-item" onclick="editarFNCSolicitante('${item.id}')"><i class="bi bi-pencil-square me-1"></i> Editar ou Converter em FNC</a>`)
                    itens.push(`<a class="dropdown-item" onclick="confirmacaoCancelamento('${item.id}')"><i class="bi bi-x-lg me-1"></i> Cancelar</a>`)
                }

            }

            itens.push(
                `<hr class="dropdown-divider">`,
                `<a class="dropdown-item" onclick="abrirModalHistoricoRespostas('${item.id}', '${item.codFilial}')"><i class="bi bi-clock-history me-1"></i> Histórico de Respostas</a>`,
                `<a class="dropdown-item" onclick="abrirModalHistoricoAprovacoes('${item.id}', '${item.codFilial}')"><i class="bi bi-clock-history me-1"></i> Histórico de Aprovações</a>`,
                `<a class="dropdown-item" onclick="abrirModalHistoricoObservacoes('${item.id}')"><i class="bi bi-clock-history me-1"></i> Histórico de Observações</a>`,
            )

            return `
                <td style="text-align: center">
                    <dropdown-menu label="Ações">
                        ${itens.join('')}
                    </dropdow-menu>
                </td>
            `;
        }

        function obterAcoesAnalista(item) {
            const itens = []

            if(item.tipoReclamacao != 'NOTIFICACAO') {
                if (item.situacao != eSituacao.ENCERRADA) {
                    itens.push(`<a class="dropdown-item" onclick="abrirModalEnvioResponsavel('${item.id}', '${item.codFilial}', '${item.responsavelId}')"><i class="bi bi-box-arrow-up-right me-1"></i> Enviar para Responsável</a>`)
                }

                itens.push(
                    `<a class="dropdown-item" onclick="editarFNCAnalista('${item.id}')"><i class="bi bi-pencil-square me-1"></i> Editar/Analisar</a>`,
                    `<a class="dropdown-item" onclick="confirmacaoEnviarParaAprovacao('${item.id}')"><i class="bi bi-box-arrow-up-right me-1"></i> Enviar para Aprovação</a>`,
                )
            }
            
            itens.push(
                `<hr class="dropdown-divider">`,
                `<a class="dropdown-item" onclick="abrirModalHistoricoRespostas('${item.id}', '${item.codFilial}')"><i class="bi bi-clock-history me-1"></i> Histórico de Respostas</a>`,
                `<a class="dropdown-item" onclick="abrirModalHistoricoAprovacoes('${item.id}', '${item.codFilial}')"><i class="bi bi-clock-history me-1"></i> Histórico de Aprovações</a>`,
                `<a class="dropdown-item" onclick="abrirModalHistoricoObservacoes('${item.id}')"><i class="bi bi-clock-history me-1"></i> Histórico de Observações</a>`,
            );

            return `
                <td style="text-align: center">
                    <dropdown-menu label="Ações">
                        ${itens.join('')}
                    </dropdow-menu>
                </td>
            `;
        }

        function obterAcoesResponsavel(item) {
            const itens = [
                `<a class="dropdown-item" onclick="editarFNCAnalista('${item.id}')"><i class="bi bi-pencil-square me-1"></i> Editar</a>`
            ]

            if($('#flagApenasPendentes').is(':checked')) {
                itens.push(
                    `<a class="dropdown-item" onclick="responderFNC('${item.id}')"><i class="bi bi-send me-1"></i> Responder</a>`,
                )
            }

            itens.push(
                `<hr class="dropdown-divider">`,
                `<a class="dropdown-item" onclick="abrirModalHistoricoRespostas('${item.id}', '${item.codFilial}')"><i class="bi bi-clock-history me-1"></i> Histórico de Respostas</a>`,
                `<a class="dropdown-item" onclick="abrirModalHistoricoObservacoes('${item.id}')"><i class="bi bi-clock-history me-1"></i> Histórico de Observações</a>`,
            )

            return `
                <td style="text-align: center">
                    <dropdown-menu label="Ações">
                        ${itens.join('')}
                    </dropdow-menu>
                </td>
            `;
        }

        function obterAcoesAprovador(item) {
            const itens = [
                `<a class="dropdown-item" onclick="editarFNCAnalista('${item.id}')"><i class="bi bi-pencil-square me-1"></i> Editar</a>`
            ]

            if($('#flagApenasPendentes').is(':checked')) {
                itens.push(
                    `<a class="dropdown-item" onclick="avaliarFNC('${item.id}')"><i class="bi bi-clipboard-data me-1"></i> Aprovar ou Rejeitar</a>`,
                )
            }

            itens.push(
                `<hr class="dropdown-divider">`,
                `<a class="dropdown-item" onclick="abrirModalHistoricoRespostas('${item.id}', '${item.codFilial}')"><i class="bi bi-clock-history me-1"></i> Histórico de Respostas</a>`,
                `<a class="dropdown-item" onclick="abrirModalHistoricoObservacoes('${item.id}')"><i class="bi bi-clock-history me-1"></i> Histórico de Observações</a>`,
                `<a class="dropdown-item" onclick="abrirModalHistoricoAprovacoes('${item.id}', '${item.codFilial}')"><i class="bi bi-clock-history me-1"></i> Histórico de Aprovações</a>`,
            )

            return `
                <td style="text-align: center">
                    <dropdown-menu label="Ações">
                        ${itens.join('')}
                    </dropdow-menu>
                </td>
            `;
        }

        function confirmacaoEnviarParaAprovacao(idFNC) {
            ConfirmationModal.warning({
                message: 'Deseja realmente enviar essa FNC para aprovação? Essa ação não poderá ser revertida',
                confirmText: 'Confirmar',
                cancelText: 'Cancelar',
                onConfirm: () => {
                    enviarParaAprovacao(idFNC);
                },
                onCancel: () => {}
            });
        }

        function confirmacaoCancelamento(idFNC) {
            ConfirmationModal.danger({
                message: 'Deseja realmente cancelar essa reclamação? Essa ação não poderá ser revertida',
                confirmText: 'Confirmar',
                cancelText: 'Cancelar',
                onConfirm: () => {
                    cancelarFNC(idFNC);
                },
                onCancel: () => {}
            });
        }
        
        function cancelarFNC(idFNC) {         
            exibirLoading(true);

            const params = {
                acao: 'cancelar',
                idFNC: idFNC
            };

            $.ajax({
                url: 'FNCRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: params,
                dataType: 'json',
                success: function(response) {
                    exibirLoading(false);

                    if (!response.sucesso) {
                        Toast.error(response.mensagem || 'Erro ao cancelar');
                        return;
                    }

                    Toast.success(response.mensagem || 'Reclamação cancelada com sucesso'); 
                    carregarFNCs();
                },
                error: function(xhr, status, error) {
                    exibirLoading(false);
                    console.error('Erro ao cancelar:', error);
                    Toast.error('Erro ao cancelar');
                }
            });
        }
        
        function enviarParaAprovacao(idFNC) {            
            const params = {
                acao: 'enviarParaAprovacao',
                idFNC: idFNC
            };

            $.ajax({
                url: 'FNCRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: params,
                dataType: 'json',
                success: function(response) {
                    if (!response.sucesso) {
                        Toast.error(response.mensagem || 'Erro ao enviar para aprovação');
                        return;
                    }

                    Toast.success(response.mensagem || 'FNC enviada para aprovação com sucesso'); 
                    carregarFNCs();
                },
                error: function(xhr, status, error) {
                    console.error('Erro ao enviar para aprovação:', error);
                    Toast.error('Erro ao enviar para aprovação');
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

        function abrirModalHistoricoRespostas(id, codPlanta) {
            $('#idHistoricoResposta').text(id);
            const modal = new bootstrap.Modal(document.getElementById('modalHistoricoRespostas'));
            modal.show();

            const $tbody = $('#gradeHistoricoResposta tbody');
            $tbody.empty();
     
            $.ajax({
                url: 'FNCRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: {
                    acao: "listarHistoricoRespostas",
                    codPlanta: codPlanta,
                    id: id
                },
                dataType: 'json',
                success: function(response) {
                    if (!response.sucesso) {
                        Toast.error(response.mensagem || 'Erro ao buscar respostas.');

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
                                <td>${item.dataInclusao || '-'}</td>
                                <td>${item.dataFinalizacao || '-'}</td>
                                <td>${item.nomeUsuarioInclusao || '-'}</td>
                                <td>${item.nomeUsuarioFinalizacao || '-'}</td>
                                <td>${item.avaliacao == '' ? 'Ainda não avaliado' : item.avaliacao == 1 ? 'Sim' : 'Não'}</td>
                                ${criarCelulaDetalhe(item.descricao, 'Descrição da resposta')}
                                ${criarCelulaDetalhe(item.justificativa, 'Justificativa da resposta')}
                            </tr>
                        `);
                    });
                },
                error: function(xhr, status, error) {
                    console.error('Erro ao buscar histórico de respostas:', error);
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

        function abrirModalHistoricoAprovacoes(id, codPlanta) {
            $('#idHistoricoAprovacao').text(id);
            const modal = new bootstrap.Modal(document.getElementById('modalHistoricoAprovacoes'));
            modal.show();

            const $tbody = $('#gradeHistoricoAprovacao tbody');
            $tbody.empty();
     
            $.ajax({
                url: 'FNCRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: {
                    acao: "listarHistoricoAprovacoes",
                    codPlanta: codPlanta,
                    id: id
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
                            <tr style="${item.invalido == '1' ? 'background-color: #f9e4e4;' : ''}">
                                <td>${item.dataInclusao || '-'}</td>
                                <td>${item.dataAvaliacao || '-'}</td>
                                <td>${item.nomeUsuarioInclusao || '-'}</td>
                                <td>${item.nomeUsuarioAvaliacao || '-'}</td>
                                <td>${item.nivel || '-'}</td>
                                <td>${item.avaliacao || '-'}</td>
                                ${criarCelulaDetalhe(item.motivo, 'Justificativa da aprovação')}
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
        
        function abrirModalHistoricoObservacoes(id, codPlanta) {
            $('#idHistoricoObservacao').text(id);
            const modal = new bootstrap.Modal(document.getElementById('modalHistoricoObservacoes'));
            modal.show();

            const $tbody = $('#gradeHistoricoObservacao tbody');
            $tbody.empty();
     
            $.ajax({
                url: 'FNCRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: {
                    acao: "listarHistoricoObservacoes",
                    codPlanta: codPlanta,
                    id: id
                },
                dataType: 'json',
                success: function(response) {
                    if (!response.sucesso) {
                        Toast.error(response.mensagem || 'Erro ao buscar observações.');

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
                                <td>${item.id || '-'}</td>
                                ${criarCelulaDetalhe(item.descricao, 'Observação')}
                                <td>${item.nomeUsuarioInclusao || '-'}</td>
                                <td>${item.dataCriacao || '-'}</td>
                            </tr>
                        `);
                    });
                },
                error: function(xhr, status, error) {
                    console.error('Erro ao buscar histórico de observações:', error);
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

        function abrirModalEnvioResponsavel(idFNC, codPlanta, responsavelId) {
            $('#idEnvioResponsavel').text(idFNC);
            const modal = new bootstrap.Modal(document.getElementById('modalEnvioResponsavel'));
            modal.show();

            carregarResponsaveis(codPlanta, responsavelId);
        }

        function enviarParaResponsavel() {
            const $responsavel = $('#responsavelIdEnvio')
            if (!$responsavel.val()) {
                Toast.error('Selecione um responsável');
                $responsavel.focus();
                return;
            }

            const $btnEnviarResponsavel = $('#btnEnviarResponsavel');
            const $loadingIcon = $btnEnviarResponsavel.find('.loading-icon');
            const $checkIcon = $btnEnviarResponsavel.find('i');
            const modal = bootstrap.Modal.getInstance(document.getElementById('modalEnvioResponsavel'));

            $checkIcon.hide();
            $btnEnviarResponsavel.prop('disabled', true);
            $loadingIcon.addClass('spinner-border spinner-border-sm').css('display', 'inline-block');

            const dadosFormulario = {
                acao:'enviarParaResponsavel',
                idFNC: $('#idEnvioResponsavel').text(),
                responsavelId: $('#responsavelIdEnvio').val()
            };

            $.ajax({
                url: 'FNCRepositorio.asp',
                type: 'POST',
                data: dadosFormulario,
                dataType: 'json',
                success: function(response) {
                    if (!response.sucesso) {
                        Toast.error(response.mensagem);
                        $checkIcon.show();
                        $btnEnviarResponsavel.prop('disabled', false);
                        $loadingIcon.removeClass('spinner-border spinner-border-sm').css('display', 'none');
                        return;
                    }

                    $checkIcon.show();
                    $btnEnviarResponsavel.prop('disabled', false);
                    $loadingIcon.removeClass('spinner-border spinner-border-sm').css('display', 'none');
                    
                    Toast.success(response.mensagem);
                    modal.hide();
                },
                error: function(xhr, status, error) {
                    Toast.error('Erro ao salvar formulário');
                    $checkIcon.show();
                    $btnEnviarResponsavel.prop('disabled', false);
                    $loadingIcon.removeClass('spinner-border spinner-border-sm').css('display', 'none');
                }
            });
            
        }

        function carregarLimitesNQA() {
            return new Promise((resolve, reject) => {
                $.ajax({
                url: 'FNCRepositorio.asp?v=<%=Now()%>',
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

            if(nivelUsuario == 'SOLICITANTE' || nivelUsuario == 'ANALISTA') {
                $('#divNovoRegistro').removeClass('d-none');
            }

            inicializarSelect2();

            const urlParams = new URLSearchParams(window.location.search);

            if (urlParams.has('id')) 
                $('#filtroIdFNC').val(urlParams.get('id'));

            if (urlParams.has('codPlanta')) 
                $('#filtroFilial').val(urlParams.get('codPlanta')).trigger('change');

            respostaNQA = await carregarLimitesNQA();
            if (respostaNQA.sucesso) limitesNQA = respostaNQA.dados;

            carregarFNCs();
        });
    </script>
</head>
<body>
    <div class="container my-4">
        <div class="my-4">
            <h3>Lista de FNC's</h3>
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
                        <option value="08">Araçatuba</option>
                        <option value="09">Blumenau</option>
                        <option value="10">Pacajus</option>
                        <option value="11">Porto Feliz</option>
                        <option value="12">Três Barras</option>
                        <option value="13">Florestal</option>
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

                <input type="hidden" id="hidNomeCliente" name="hidNomeCliente">
                
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
                        <option value="1">Registrada</option>
                        <option value="2">Em Análise</option>
                        <option value="3">Procede</option>
                        <option value="4">Não Procede</option>
                        <option value="5">Cancelada</option>
                    </select>
                </div>
                
                <div class="col-12 col-md-4 col-lg-3">
                    <label for="filtroAprovada" class="form-label">Aprovada</label>
                    <select required class="form-select form-select-sm select2" id="filtroAprovada" name="filtroAprovada" data-placeholder="Todos">
                        <option value="">Todos</option>
                        <option value="1">Sim</option>
                        <option value="0">Não</option>
                    </select>
                </div>

                <div class="col-12 col-md-4 col-lg-3">
                    <label for="filtroSituacao" class="form-label">Situação</label>
                    <select required class="form-select form-select-sm select2" id="filtroSituacao" name="filtroSituacao" data-placeholder="Todos">
                        <option value="">Todos</option>
                        <option value="1">Aberta</option>
                        <option value="2">Respondida</option>
                        <option value="3">Encerrada</option>
                    </select>
                </div>

                 <div class="col-12 col-md-4 col-lg-3">
                    <label for="filtroTipoReclamacao" class="form-label">Tipo de Reclamação</label>
                    <select required class="form-select form-select-sm select2" id="filtroTipoReclamacao" name="filtroTipoReclamacao" data-placeholder="Todos">
                        <option value="">Todos</option>
                        <option value="NOTIFICACAO">NOTIFICACAO</option>
                        <option value="FNC">FNC</option>
                    </select>
                </div>

                <div class="col-12 col-md-4 col-lg-3">
                    <label for="filtroIdFNC" class="form-label">Id FNC</label>
                    <input type="number" class="form-control form-control-sm" id="filtroIdFNC" name="filtroIdFNC" placeholder="Id da FNC">
                </div>

                <div class="col-12 mt-4 d-flex justify-content-between">
                    <div class="d-flex align-items-end">
                        <div class="form-check form-switch">
                            <input onchange="carregarFNCs()" checked class="form-check-input" type="checkbox" name="flagApenasPendentes" id="flagApenasPendentes" />
                            <label class="form-check-label" for="flagApenasPendentes">Apenas minhas pendências</label>
                        </div>
                    </div>

                    <div class="d-flex justify-content-end align-items-end gap-2">
                        <button onclick="limparFiltros()" type="button" class="btn btn-sm btn-outline-secondary" id="btnLimpar">
                            <i class="bi bi-x-lg me-1"></i> Limpar
                        </button>
                        <button onclick="carregarFNCs()" type="button" class="btn btn-sm btn-primary" id="btnPesquisar">
                            <i class="bi bi-search me-1"></i> Pesquisar
                        </button>
                    </div>
                </div>
            </div>
        </div>

        <div class="alert alert-primary alert-dismissible py-2 px-3 mb-4 d-flex gap-1 align-items-center" role="alert" style="font-size: 0.85rem;">
            <i class="bi bi-info-circle me-2"></i>
            Visualizando como: <span id="nivelUsuario"></span>
        </div>

        <div class="card">
            <div id="divNovoRegistro" class="d-none gap-2 p-3 border-bottom">
                <button class="btn btn-sm btn-primary" onclick="novaFNC()">
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

                <table class="table table-hover mb-0" id="gradeFNC">
                    <thead>
                        <tr>
                            <th>#</th>
                            <th>Data Inclusão</th>
                            <th>Filial</th>
                            <th>Tipo de Reclamação</th>
                            <th>Setor</th>
                            <th>Responsável</th>
                            <th>Cliente</th>
                            <th>Loja</th>
                            <th>Quantidade</th>
                            <th>Disposição</th>
                            <th>Defeito</th>
                            <th>Aprovada</th>
                            <th>Coeficiente NQA</th>
                            <th>Status</th>
                            <th>Situação</th>
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

    <!-- Modal de Histórico de respostas -->
    <div class="modal fade" id="modalHistoricoRespostas" tabindex="-1" aria-labelledby="modalHistoricoRespostasLabel" aria-hidden="true">
        <div class="modal-dialog modal-xl">
            <div class="modal-content">
                <div class="modal-header">
                    <h5 class="modal-title" id="modalHistoricoRespostasTitulo">
                        Histórico de Respostas #<span id="idHistoricoResposta"></span>
                    </h5>
                    <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Fechar"></button>
                </div>
                <div class="modal-body">
                    <table class="table table-hover mb-0" id="gradeHistoricoResposta">
                        <thead>
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

    <!-- Modal de Histórico de aprovações -->
    <div class="modal fade" id="modalHistoricoAprovacoes" tabindex="-1" aria-labelledby="modalHistoricoAprovacoesLabel" aria-hidden="true">
        <div class="modal-dialog modal-xl">
            <div class="modal-content">
                <div class="modal-header">
                    <h5 class="modal-title" id="modalHistoricoAprovacoesTitulo">
                        Histórico de Aprovações #<span id="idHistoricoAprovacao"></span>
                    </h5>
                    <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Fechar"></button>
                </div>
                <div class="modal-body">
                    <table class="table table-hover mb-0" id="gradeHistoricoAprovacao">
                        <thead>
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

    <!-- Modal de Histórico de observações -->
    <div class="modal fade" id="modalHistoricoObservacoes" tabindex="-1" aria-labelledby="modalHistoricoObservacoesLabel" aria-hidden="true">
        <div class="modal-dialog modal-xl">
            <div class="modal-content">
                <div class="modal-header">
                    <h5 class="modal-title" id="modalHistoricoObservacoesTitulo">
                        Histórico de Observações #<span id="idHistoricoObservacao"></span>
                    </h5>
                    <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Fechar"></button>
                </div>
                <div class="modal-body">
                    <table class="table table-hover mb-0" id="gradeHistoricoObservacao">
                        <thead>
                            <tr>
                                <th>#</th>
                                <th>Descrição</th>
                                <th>Incluida Por</th>
                                <th>Data Inclusão</th>
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

    <!-- Modal de Envio para responsável -->
     <div class="modal fade" id="modalEnvioResponsavel" tabindex="-1" aria-labelledby="modalEnvioResponsavelLabel" aria-hidden="true">
        <div class="modal-dialog modal-xl">
            <div class="modal-content">
                <div class="modal-header">
                        <h5 class="modal-title" id="modalEnvioResponsavelTitulo">
                        Enviar para Responsável #<span id="idEnvioResponsavel"></span>
                    </h5>
                    <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Fechar"></button>
                </div>

                <div class="modal-body">
                    <div class="row mb-3">
                        <div class="col-12">
                            <label for="responsavelIdEnvio" class="form-label">Responsável <span class="required-field">*</span></label>
                            <select required class="form-select form-select-sm select2" name="responsavelIdEnvio" id="responsavelIdEnvio">
                                <option value="">Selecione um responsável</option>
                            </select>
                        </div>
                    </div>
                </div>

                <div class="modal-footer">
                    <button type="button" class="btn btn-sm btn-secondary" data-bs-dismiss="modal">Cancelar</button>
                    <button onclick="enviarParaResponsavel()" class="btn btn-sm btn-success" id="btnEnviarResponsavel">
                        <span class="loading-icon"></span>
                        <i class="bi bi-check-lg me-1"></i>
                        <span>Salvar</span>
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