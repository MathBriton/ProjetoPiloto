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
    <title>Alteração de Dados</title>
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

        #gradeSolicitacao th:last-child,
        #gradeSolicitacao td:last-child {
            position: sticky;
            right: 0;
            z-index: 5;
            white-space: nowrap;
        }
        
        #gradeSolicitacao td:last-child {
            background: linear-gradient(to right, rgba(0,0,0,0.1) 0%, transparent 1px),#fff;
            background-origin: border-box;
            background-clip: padding-box;
        }
        
        #gradeSolicitacao thead th:last-child {
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
    </style>

    <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.0.2/dist/js/bootstrap.bundle.min.js" integrity="sha384-MrcW6ZMFYlzcLA8Nl+NtUVF0sA7MsXsP1UyJoMp4YLEuNSfAP+JcXn/tWtIaxVXM" crossorigin="anonymous"></script>
    <script src="https://cdn.jsdelivr.net/npm/select2@4.1.0-rc.0/dist/js/select2.min.js"></script>
    <script src="/corporativo/components/toast.js?v=<%=Now()%>"></script>
    <script src="/corporativo/components/confirmation.js?v=<%=Now()%>"></script>
    <script src="/corporativo/funcoes/funcoesJAVASCRIPT.js?v=<%=Now()%>"></script>
    <script src="/corporativo/components/dropdownMenu.js?v=<%=Now()%>"></script>

    <script>
        let nivelUsuario = '';
        let dadosSolicitacoes = [];
        let carregando = false;
        let acaoSolicitacaoAtual = null;

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
                    url: 'SolicitacaoClienteAlteracaoRepositorio.asp?v=<%=Now()%>',
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

        function carregarSolicitacoes() { 
            exibirLoading(true);

            const params = {
                acao: 'listar',
                codigoPlanta: $('#codigoPlanta').val(),
                dataInclusaoDe: $('#dataInclusaoDe').val(),
                dataInclusaoAte: $('#dataInclusaoAte').val(),
                status: $('#status').val(),
                tipoSolicitacao: $('#tipoSolicitacao').val(),
                codigoCliente: $('#codigoCliente').val(),
                flagApenasPendentes: $('#flagApenasPendentes').is(':checked') ? 1 : 0
            };

            $.ajax({
                url: 'SolicitacaoClienteAlteracaoRepositorio.asp?v=<%=Now()%>',
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
                const acoesMap = {
                    'SOLICITANTE': obterAcoesSolicitante(item),
                    'DIRETOR_COMERCIAL': obterAcoesDiretorComercial(item),
                    'CDAG': obterAcoesCDAG(item),
                    'FINANCEIRO': obterAcoesFinanceiro(item),
                    'CONTROLADORIA': obterAcoesControladoria(item)
                }

                let htmlAcoes = acoesMap[nivelUsuario] || '';
                
                $tbody.append(`
                    <tr>
                        <td>${item.id || '-'}</td>
                        <td>${formatarParaDataBR(item.dataInclusao)}</td>
                        <td>${obterBadgeTipoSolicitacao(item.tipoSolicitacao)}</td>
                        <td>${obterNomePlanta(item.codigoPlanta)}</td>
                        <td>${obterBadgeStatus(item.status)}</td>
                        <td>${item.razaoSocial}</td>
                        <td>${item.tipoPessoa == 'F' ? formatarCPF(item.cpfCnpj) : formatarCNPJ(item.cpfCnpj)}</td>
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

        function novaSolicitacao() {
            window.location.href = `SolicitacaoClienteAlteracaoFormularioGeral.asp?acao=CRIAR`;
        }

        function editarSolicitacao(id) {
            window.location.href = `SolicitacaoClienteAlteracaoFormularioGeral.asp?acao=EDITAR&id=${id}`;
        }

        function visualizarSolicitacao(id) {
            window.location.href = `SolicitacaoClienteAlteracaoFormularioGeral.asp?acao=VISUALIZAR&id=${id}`;
        }

        function abrirModalAcao(configuracao) {
            acaoSolicitacaoAtual = configuracao;

            $('#idAcao').text(configuracao.id);
            $('#modalAcaoTituloTexto').text(configuracao.titulo);
            $('#observacao').val('');
            $('#observacao').removeClass('is-invalid');
            $('#feedbackObservacao').text('');
            $('#labelObservacaoObrigatoria').text(configuracao.observacaoObrigatoria ? 'obrigatória' : 'opcional');

            const modal = new bootstrap.Modal(document.getElementById('modalAcao'));
            modal.show();
        }

        function abrirModalEnviarParaFinanceiro(id) {
            abrirModalAcao({
                acao: 'enviarParaFinanceiro',
                id: id,
                titulo: 'Enviar para Financeiro',
                observacaoObrigatoria: false,
            });
        }

        function abrirModalEnviarParaDiretorComercial(id) {
            abrirModalAcao({
                acao: 'enviarParaDiretorComercial',
                id: id,
                titulo: 'Enviar para Diretor Comercial',
                observacaoObrigatoria: false,
            });
        }

        function abrirModalEnviarParaControladoria(id) {
            abrirModalAcao({
                acao: 'enviarParaControladoria',
                id: id,
                titulo: 'Enviar para Controladoria',
                observacaoObrigatoria: false,
            });
        }

        function abrirModalCancelar(id) {
            abrirModalAcao({
                acao: 'cancelar',
                id: id,
                titulo: 'Cancelar Solicitação',
                observacaoObrigatoria: true,
            });
        }

        function abrirModalFinalizar(id) {
            abrirModalAcao({
                acao: 'finalizar',
                id: id,
                titulo: 'Finalizar Solicitação',
                observacaoObrigatoria: false,
            });
        }

        function abrirModalHistorico(id) {
            $('#idHistorico').text(id);
            const modal = new bootstrap.Modal(document.getElementById('modalHistorico'));
            modal.show();

            const $tbody = $('#gradeHistorico tbody');
            $tbody.empty();
     
            $.ajax({
                url: 'SolicitacaoClienteAlteracaoRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: {
                    acao: "listarHistorico",
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
                                <td>${item.dataInclusao || '-'}</td>
                                <td>${item.loginUsuarioInclusao || '-'}</td>
                                <td>${item.acao || '-'}</td>
                                ${criarCelulaDetalhe(item.observacao, 'Observação ')}
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
    
        function abrirModalDetalhe(botao) {
            const titulo = $(botao).attr('data-titulo') || 'Detalhes';
            const detalhe = $(botao).attr('data-detalhe') || '-';

            $('#modalDetalheTituloTexto').text(titulo);
            $('#modalDetalheConteudo').text(detalhe);

            const modal = new bootstrap.Modal(document.getElementById('modalDetalheTexto'));
            modal.show();
        }

        function abrirFormularioAvaliacao(id, acao) {
            window.location.href = `SolicitacaoClienteAlteracaoFormularioAvaliacao.asp?id=${id}&acao=${acao}`;
        }

        function executarAcao() {
            if (!acaoSolicitacaoAtual) {
                return;
            }

            const modal = bootstrap.Modal.getInstance(document.getElementById('modalAcao'));
            const observacao = $.trim($('#observacao').val());

            $('#observacao').removeClass('is-invalid');
            $('#feedbackObservacao').text('');

            if (acaoSolicitacaoAtual.observacaoObrigatoria && !observacao) {
                $('#observacao').addClass('is-invalid');
                $('#feedbackObservacao').text('Informe uma observação para continuar.');
                return;
            }

            modal.hide();
            exibirLoading(true);

            const params = {
                acao: acaoSolicitacaoAtual.acao,
                id: acaoSolicitacaoAtual.id,
                observacao: observacao
            };

            $.ajax({
                url: 'SolicitacaoClienteAlteracaoRepositorio.asp?v=<%=Now()%>',
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
                    acaoSolicitacaoAtual = null;

                    carregarSolicitacoes();
                },
                error: function(xhr, status, error) {
                    exibirLoading(false);
                    console.error(error);
                    Toast.error(error);
                }
            });
        }

        function limparFiltros() {
            $('#codigoPlanta').val('').trigger('change');
            $('#dataInclusaoDe').val('');
            $('#dataInclusaoAte').val('');
            $('#codigoCliente').val('');
            $('#status').val('').trigger('change');

            dadosSolicitacoes = [];
            renderizarGrade();
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
       
        function obterBadgeTipoSolicitacao(tipoSolicitacao) {
            switch(String(tipoSolicitacao || '').trim().toUpperCase()) {
                case 'BLOQUEIO':
                    return `<span class="badge badge-red"><i class="bi bi-lock-fill me-1"></i>Bloqueio Simples</span>`;
                case 'DESBLOQUEIO':
                    return `<span class="badge badge-green"><i class="bi bi-unlock-fill me-1"></i>Desbloqueio Simples</span>`;
                case 'DESBLOQUEIO_COM_ALTERACAO_DADOS':
                    return `<span class="badge badge-green"><i class="bi bi-unlock-fill me-1"></i>Desbloqueio c/ Alteração de Dados</span>`;
                case 'ALTERACAO_DADOS':
                    return `<span class="badge badge-blue"><i class="bi bi-repeat me-1"></i>Alteração de Dados</span>`;
                default:
                    return tipoSolicitacao || '-';
            }
        }

        function obterBadgeStatus(status) {
            switch(status) {
                case 'EM_DIGITACAO':
                    return `<span class="badge badge-yellow">Em Digitação</span>`;
                case 'PENDENTE_DIRETOR_COMERCIAL':
                    return `<span class="badge badge-blue">Pendente Diretor</span>`;
                case 'PENDENTE_FINANCEIRO':
                    return `<span class="badge badge-purple">Pendente Financeiro</span>`;
                case 'PENDENTE_CONTROLADORIA':
                    return `<span class="badge badge-teal">Pendente Controladoria</span>`;
                case 'PENDENTE_FINALIZACAO':
                    return `<span class="badge badge-teal">Pendente Finalização</span>`;
                case 'FINALIZADO':
                    return `<span class="badge badge-green">Finalizado</span>`;
                case 'CANCELADO':
                    return `<span class="badge badge-gray">Cancelado</span>`;
                default:
                    return '-';
            }
        }

        function obterItemAcao(props, icone, texto, evento) {
            return `<a class="dropdown-item" onclick="${evento}('${props.id}', '${props.acao}')"><i class="${icone} me-1"></i> ${texto}</a>`;
        }

        function obterMenuAcoes(itens) {
            return `
                <td style="text-align: center">
                    <dropdown-menu label="Ações">
                        ${itens.join('')}
                    </dropdow-menu>
                </td>
            `
        }

        function adicionarAcaoHistorico(itens, id) {
            itens.push(
                `<hr class="dropdown-divider">`,
                obterItemAcao({id}, 'bi bi-clock-history', 'Histórico', 'abrirModalHistorico')
            );
        }

        function obterAcoesSolicitante(item) {
            const id = item.id;
            const status = item.status;
            const itens = [
                obterItemAcao({id}, 'bi bi-eye', 'Visualizar', 'visualizarSolicitacao'),
            ];

            if (status == 'EM_DIGITACAO') {
                itens.push(obterItemAcao({id}, 'bi bi-pencil-square', 'Editar', 'editarSolicitacao'));
                itens.push(obterItemAcao({id}, 'bi bi-ban', 'Cancelar', 'abrirModalCancelar'));

                // Situações listadas abaixo precisam passar pela aprovação do comercial
                // Outras situações vão direto para financeiro
                const deveEnviarParaComercial = 
                    (item.tipoSolicitacao == 'ALTERACAO_DADOS' && item.possuiCondicaoPagamento == 1) ||
                    (item.tipoSolicitacao == 'DESBLOQUEIO_COM_ALTERACAO_DADOS' && item.possuiCodigoVendedor == 1) ||
                    (item.tipoSolicitacao == 'BLOQUEIO')   
                
                // 
                const deveEnviarParaControladoria = 
                    (item.tipoSolicitacao == 'ALTERACAO_DADOS' && item.possuiCondicaoPagamento == 0) 

                if(deveEnviarParaComercial) {
                    itens.push(obterItemAcao({id}, 'bi bi-box-arrow-up-right', 'Enviar Para Diretor Comercial', 'abrirModalEnviarParaDiretorComercial'))
                } else if(deveEnviarParaControladoria) {
                    itens.push(obterItemAcao({id}, 'bi bi-box-arrow-up-right', 'Enviar Para Controladoria', 'abrirModalEnviarParaControladoria'))
                } else {
                    itens.push(obterItemAcao({id}, 'bi bi-box-arrow-up-right', 'Enviar Para Financeiro', 'abrirModalEnviarParaFinanceiro'))
                }
            }

            adicionarAcaoHistorico(itens, id);
            return obterMenuAcoes(itens);
        }

        function obterAcoesDiretorComercial(item) {
            const id = item.id;
            const status = item.status;
            const itens = [
                obterItemAcao({id}, 'bi bi-eye', 'Visualizar', 'visualizarSolicitacao'),
            ];

            if (status == 'PENDENTE_DIRETOR_COMERCIAL') {
                itens.push(
                    obterItemAcao({id, acao: 'AVALIACAO_DIRETOR_COMERCIAL'}, 'bi bi-box-arrow-up-right', 'Avaliar', 'abrirFormularioAvaliacao'),
                );
            }

            adicionarAcaoHistorico(itens, id);
            return obterMenuAcoes(itens);
        }

        function obterAcoesCDAG(item) {
            const id = item.id;
            const status = item.status;
            const itens = [
                obterItemAcao({id}, 'bi bi-eye', 'Visualizar', 'visualizarSolicitacao')
            ];

            if (status == 'PENDENTE_FINALIZACAO') {
                itens.push(
                    obterItemAcao({id}, 'bi bi-check-circle', 'Marcar como Finalizado', 'abrirModalFinalizar'),
                );
            }

            adicionarAcaoHistorico(itens, id);
            return obterMenuAcoes(itens);
        }
        
        function obterAcoesFinanceiro(item) {
            const id = item.id;
            const status = item.status;
            const itens = [
                obterItemAcao({id}, 'bi bi-eye', 'Visualizar', 'visualizarSolicitacao'),
            ];

            if (status == 'PENDENTE_FINANCEIRO') {
                itens.push(
                    obterItemAcao({id, acao: 'AVALIACAO_FINANCEIRO'}, 'bi bi-box-arrow-up-right', 'Avaliar', 'abrirFormularioAvaliacao'),
                );
            }

            adicionarAcaoHistorico(itens, id);
            return obterMenuAcoes(itens);
        }
      
        function obterAcoesControladoria(item) {
            const id = item.id;
            const status = item.status;
            const itens = [
                obterItemAcao({id}, 'bi bi-eye', 'Visualizar', 'visualizarSolicitacao'),
            ];

            if (status == 'PENDENTE_CONTROLADORIA') {
                itens.push(
                    obterItemAcao({id, acao: 'AVALIACAO_CONTROLADORIA'}, 'bi bi-box-arrow-up-right', 'Avaliar', 'abrirFormularioAvaliacao'),
                );
            }

            adicionarAcaoHistorico(itens, id);
            return obterMenuAcoes(itens);
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

            if(nivelUsuario == 'SOLICITANTE') {
                $('#btnNovaSolicitacao').removeClass('d-none');
            }

            inicializarSelect2();
            carregarSolicitacoes();
        });
    </script>
</head>
<body>
    <div class="container my-4">
        <div class="my-4">
            <h3>Solicitações de Alteração de Dados Cliente</h3>
        </div>

        <div class="card p-3 mb-4">
            <div class="row g-3">
                <div class="col-12 col-md-4 col-lg-3">
                    <label for="codigoPlanta" class="form-label">Filial</label>
                    <select required class="form-select form-select-sm select2" id="codigoPlanta" name="codigoPlanta" data-placeholder="Todas as filiais">
                        <option value="">Todas as filiais</option>
                        <option value="00">Pirapetinga</option>
                        <option value="04">Uberaba</option>
                        <option value="05">Bento</option>
                        <option value="06">Maranguape</option>
                        <option value="07">Saquarema</option>
                    </select>
                </div>

                <div class="col-12 col-md-4 col-lg-3">
                    <label for="dataInclusaoDe" class="form-label">Data Inclusão De</label>
                    <input type="date" class="form-control form-control-sm" id="dataInclusaoDe" name="dataInclusaoDe">
                </div>

                <div class="col-12 col-md-4 col-lg-3">
                    <label for="dataInclusaoAte" class="form-label">Data Inclusão Até</label>
                    <input type="date" class="form-control form-control-sm" id="dataInclusaoAte" name="dataInclusaoAte">
                </div>
                
                <div class="col-12 col-md-4 col-lg-3">
                    <label for="codigoCliente" class="form-label">Cód. Cliente</label>
                    <input type="text" class="form-control form-control-sm" id="codigoCliente" name="codigoCliente" placeholder="Código do cliente">
                </div>
                
                <div class="col-12 col-md-4 col-lg-3">
                    <label for="status" class="form-label">Status</label>
                    <select required class="form-select form-select-sm select2" id="status" name="status" data-placeholder="Todos">
                        <option value="">Todos</option>
                        <option value="EM_DIGITACAO">Em Digitação</option>
                        <option value="PENDENTE_DIRETOR_COMECIAL">Pendente Diretor Comercial</option>
                        <option value="PENDENTE_FINANCEIRO">Pendente Financeiro</option>
                        <option value="PENDENTE_CONTROLADORIA">Pendente Controladoria</option>
                        <option value="PENDENTE_FINALIZACAO">Pendente Finalização</option>
                        <option value="FINALIZADO">Finalizado</option>
                    </select>
                </div>
                
                <div class="col-12 col-md-4 col-lg-3">
                    <label for="tipoSolicitacao" class="form-label">Tipo de Solicitação</label>
                    <select required class="form-select form-select-sm select2" id="tipoSolicitacao" name="tipoSolicitacao" data-placeholder="Todos">
                        <option value="">Todos</option>
                        <option value="DESBLOQUEIO">Desbloqueio Simples</option>
                        <option value="DESBLOQUEIO_COM_ALTERACAO_DADOS">Desbloqueio c/ Alteração de Dados</option>
                        <option value="BLOQUEIO">Bloqueio Simples</option>
                        <option value="ALTERACAO_DADOS">Alteração de Dados</option>
                    </select>
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

        <div class="alert alert-primary alert-dismissible py-2 px-3 mb-4 d-flex gap-1 align-items-center" role="alert" style="font-size: 0.85rem;">
            <i class="bi bi-info-circle me-2"></i>
            Visualizando como: <span id="nivelUsuario"></span>
        </div>

        <div class="card ">
            <div class="gap-2 p-3 border-bottom">
                <button id="btnNovaSolicitacao" class="d-none btn btn-sm btn-primary" onclick="novaSolicitacao()">
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
                            <th>Data Inclusão</th>
                            <th>Tipo de Solicitação</th>
                            <th>Planta</th>
                            <th>Status</th>
                            <th>Cliente</th>
                            <th>CNPJ</th>
                            <th>Emitente</th>
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

    <!-- Modal de Histórico de aprovações -->
    <div class="modal fade" id="modalHistorico" tabindex="-1" aria-labelledby="modalHistoricoLabel" aria-hidden="true">
        <div class="modal-dialog modal-xl">
            <div class="modal-content">
                <div class="modal-header">
                    <h5 class="modal-title" id="modalHistoricoTitulo">
                        Histórico de Solicitação #<span id="idHistorico"></span>
                    </h5>
                    <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Fechar"></button>
                </div>
                <div class="modal-body">
                    <table class="table table-hover mb-0" id="gradeHistorico">
                        <thead>
                            <tr>
                                <th>#</th>
                                <th>Data</th>
                                <th>Usuário</th>
                                <th>Ação</th>
                                <th>Observação</th>
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
    
    <!-- Modal de confirmação de ação -->
    <div class="modal fade" id="modalAcao" tabindex="-1" aria-labelledby="modalAcaoLabel" aria-hidden="true">
        <div class="modal-dialog modal-xl">
            <div class="modal-content">
                <div class="modal-header">
                        <h5 class="modal-title" id="modalAcaoTitulo">
                        <span id="modalAcaoTituloTexto">Confirmar ação</span> #<span id="idAcao"></span>
                    </h5>
                    <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Fechar"></button>
                </div>

                <div class="modal-body">
                    <div class="row mb-3">
                        <div class="col-12">
                            <label for="observacao" class="form-label"> Observações (<span id="labelObservacaoObrigatoria">opcional</span>) </label>
                            <textarea class="form-control form-control-sm" id="observacao" rows="5"></textarea>
                            <div class="invalid-feedback" id="feedbackObservacao"></div>
                        </div>
                    </div>
                </div>

                <div class="modal-footer">
                    <button type="button" class="btn btn-sm btn-outline-secondary" data-bs-dismiss="modal">Cancelar</button>
                    <button onclick="executarAcao()" class="btn btn-sm btn-success" id="btnEnviarResponsavel">
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