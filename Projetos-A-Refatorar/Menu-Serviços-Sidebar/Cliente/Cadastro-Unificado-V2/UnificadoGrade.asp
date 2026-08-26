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
    <title>Cadastro Unificado</title>
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

        #gradeCadastro th,
        #gradeCadastro td {
            white-space: nowrap;
            vertical-align: top;
        }

        #gradeCadastro th:last-child,
        #gradeCadastro td:last-child {
            position: sticky;
            right: 0;
            z-index: 5;
            white-space: nowrap;
        }
        
        #gradeCadastro td:last-child {
            background: linear-gradient(to right, rgba(0,0,0,0.1) 0%, transparent 1px),#fff;
            background-origin: border-box;
            background-clip: padding-box;
        }
        
        #gradeCadastro thead th:last-child {
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
        const eStatus = Object.freeze({
            PENDENTE: 'P',
            FISCAL: 'S',
            CONTROLADORIA: 'C', 
            CDAG: 'E',
            CONTROLES_INTERNOS: 'I',
            FINALIZADO: 'F',
            DEVOLVIDO: 'R',
            CANCELADO: 'X',
            CLIENTE_PROSPECT: 'Z',
        });

        let nivelUsuario = '';
        let dadosCadastros = [];
        let carregando = false;
        let acaoCadastroAtual = null;

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
                    url: 'UnificadoRepositorio.asp?v=<%=Now()%>',
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

        function carregarCadastros() { 
            exibirLoading(true);

            const params = {
                acao: 'listar',
                filtroFilial: $('#filtroFilial').val(),
                filtroDataInclusaoDe: $('#filtroDataInclusaoDe').val(),
                filtroDataInclusaoAte: $('#filtroDataInclusaoAte').val(),
                filtroStatus: $('#filtroStatus').val(),
                filtroSituacao: $('#filtroSituacao').val(),
                filtroCodCliente: $('#filtroCodCliente').val(),
                flagApenasPendentes: $('#flagApenasPendentes').is(':checked') ? 1 : 0
            };

            $.ajax({
                url: 'UnificadoRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: params,
                dataType: 'json',
                success: function(response) {
                    exibirLoading(false);

                    if (!response.sucesso) {
                        Toast.error(response.mensagem);
                        return;
                    }
                
                    dadosCadastros = response.dados || [];
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
            const $tbody = $('#gradeCadastro tbody');
            
            $tbody.empty();
            
            if (dadosCadastros.length === 0) {
                $tbody.html(`
                    <tr>
                        <td colspan="28" class="text-center text-muted py-4">
                            Nenhum registro encontrado
                        </td>
                    </tr>
                `);
                return;
            }

            dadosCadastros.forEach(function(item) {    
                const statusBadge = obterBadgeStatus(item.status);

                const acoesMap = {
                    'SOLICITANTE': obterAcoesSolicitante(item),
                    'CDAG': obterAcoesCDAG(item),
                    'CONTROLES_INTERNOS': obterAcoesControlesInternos(item),
                    'CONTROLADORIA': obterAcoesControladoria(item),
                    'FISCAL': obterAcoesFiscal(item),
                }

                let htmlAcoes = acoesMap[nivelUsuario] || '';

                $tbody.append(`
                    <tr>
                        <td>${item.codigo || '-'}</td>
                        <td>${obterNomePlanta(item.codigoPlanta)}</td>
                        <td>${formatarParaDataBR(item.dataInclusao)}</td>
                        <td>${statusBadge}</td>
                        <td>${item.razaoSocial}</td>
                        <td>${item.tipoPessoa == 'F' ? formatarCPF(item.cpf) : formatarCNPJ(item.cnpj)}</td>
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

        function novoCadastro() {
            window.location.href = `UnificadoFormularioGeral.asp?acao=CRIAR`;
        }

        function visualizarCadastro(codigo) {
            window.location.href = `UnificadoFormularioGeral.asp?acao=VISUALIZAR&codigo=${codigo}`;
        }

        function editarCadastro(codigo) {
            window.location.href = `UnificadoFormularioGeral.asp?acao=EDITAR&codigo=${codigo}`;
        }

        function abrirModalAcaoCadastro(configuracao) {
            acaoCadastroAtual = configuracao;

            $('#idAcaoCadastro').text(configuracao.codigo);
            $('#modalAcaoCadastroTituloTexto').text(configuracao.titulo);
            $('#observacao').val('');
            $('#observacao').removeClass('is-invalid');
            $('#feedbackObservacao').text('');
            $('#labelObservacaoObrigatoria').text(configuracao.observacaoObrigatoria ? 'obrigatória' : 'opcional');

            const modal = new bootstrap.Modal(document.getElementById('modalAcaoCadastro'));
            modal.show();
        }

        function abrirModalCancelarCadastro(codigo) {
            abrirModalAcaoCadastro({
                acao: 'cancelar',
                codigo: codigo,
                titulo: 'Cancelar',
                observacaoObrigatoria: true,
                mensagemErro: 'Erro cancelar',
                mensagemSucesso: 'Cadastro cancelado com sucesso'
            });
        }

        function confirmacaoCancelarCadastro(codigo) {
            abrirModalCancelarCadastro(codigo);
        }

        function abrirModalDevolverParaRequisitante(codigo) {
            abrirModalAcaoCadastro({
                acao: 'devolver',
                codigo: codigo,
                titulo: 'Devolver para Requisitante',
                observacaoObrigatoria: true,
                mensagemErro: 'Erro devolver para requisitante',
                mensagemSucesso: 'Cadastro devolvido para requisitante com sucesso'
            });
        }

        function abrirModalEnviarParaCDAG(codigo) {
            abrirModalAcaoCadastro({
                acao: 'enviarParaCDAG',
                codigo: codigo,
                titulo: 'Enviar para CDAG',
                observacaoObrigatoria: false,
                mensagemErro: 'Erro enviar para CDAG',
                mensagemSucesso: 'Cadastro enviado para CDAG com sucesso'
            });
        }

        function abrirModalEnviarParaFiscal(codigo) {
            abrirModalAcaoCadastro({
                acao: 'enviarParaFiscal',
                codigo: codigo,
                titulo: 'Enviar para Fiscal',
                observacaoObrigatoria: false,
                mensagemErro: 'Erro enviar para fiscal',
                mensagemSucesso: 'Cadastro enviado com sucesso'
            });
        }

        function abrirModalEnviarParaControlesInternos(codigo) {
            abrirModalAcaoCadastro({
                acao: 'enviarParaControlesInternos',
                codigo: codigo,
                titulo: 'Enviar para Controles Internos',
                observacaoObrigatoria: true,
                mensagemErro: 'Erro enviar para controles internos',
                mensagemSucesso: 'Cadastro enviado com sucesso'
            });
        }

        function abrirModalHistoricoVistos(codigo) {
            $('#codigoHistoricoVistos').text(codigo);
            const modal = new bootstrap.Modal(document.getElementById('modalHistoricoVistos'));
            modal.show();

            const $tbody = $('#gradeHistoricoVistos tbody');
            $tbody.empty();
            $('#totalLeadtimeVistos').text('0');
     
            $.ajax({
                url: 'UnificadoRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: {
                    acao: "listarHistoricoVistos",
                    codigo: codigo
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

                    let totalLeadtime = 0;

                    response.dados.forEach(function(item) {   
                        const leadtime = parseFloat(item.leadtime) || 0;
                        totalLeadtime += leadtime;

                        $tbody.append(`
                            <tr>
                                <td>${item.id || '-'}</td>
                                <td>${item.dataInclusao || '-'}</td>
                                <td>${item.leadtime || 0}</td>
                                <td>${item.loginUsuarioInclusao || '-'}</td>
                                <td>${obterStatusTraduzido(item.status)}</td>
                                ${criarCelulaDetalhe(item.observacao, 'Observação do visto')}
                            </tr>
                        `);
                    });

                    $('#totalLeadtimeVistos').text(totalLeadtime);
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

        function abrirFormularioAnexo(codigo) {
            window.location.href = `UnificadoFormularioAnexo.asp?codigo=${codigo}`;
        }

        function abrirFormularioEnviarParaControladoria(codigo) {
            window.location.href = `UnificadoFormularioEnviarParaControladoria.asp?codigo=${codigo}`;
        }

        function abrirFormularioFinalizar(codigo) {
            window.location.href = `UnificadoFormularioFinalizar.asp?codigo=${codigo}`;
        }

        /* function abrirFormularioAprovarLimiteCredito(codigo) {
            window.location.href = `UnificadoFormularioAprovarLimiteCredito.asp?codigo=${codigo}`;
        } */

        function executarAcaoCadastro() {
            if (!acaoCadastroAtual) {
                return;
            }

            const modal = bootstrap.Modal.getInstance(document.getElementById('modalAcaoCadastro'));
            const observacao = $.trim($('#observacao').val());

            $('#observacao').removeClass('is-invalid');
            $('#feedbackObservacao').text('');

            if (acaoCadastroAtual.observacaoObrigatoria && !observacao) {
                $('#observacao').addClass('is-invalid');
                $('#feedbackObservacao').text('Informe uma observação para continuar.');
                return;
            }

            modal.hide();
            exibirLoading(true);

            const params = {
                acao: acaoCadastroAtual.acao,
                codigo: acaoCadastroAtual.codigo,
                observacao: observacao
            };

            $.ajax({
                url: 'UnificadoRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: params,
                dataType: 'json',
                success: function(response) {
                    exibirLoading(false);

                    if (!response.sucesso) {
                        Toast.error(response.mensagem || acaoCadastroAtual.mensagemErro);
                        return;
                    }

                    Toast.success(response.mensagem || acaoCadastroAtual.mensagemSucesso); 
                    acaoCadastroAtual = null;

                    carregarCadastros();
                },
                error: function(xhr, status, error) {
                    exibirLoading(false);
                    console.error(acaoCadastroAtual.mensagemErro + ':', error);
                    Toast.error(acaoCadastroAtual.mensagemErro);
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

            dadosCadastros = [];
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

        function obterStatusTraduzido(status) {
            switch(status) {
                case eStatus.PENDENTE:
                    return 'Criação';
                case eStatus.FISCAL:
                    return 'Enviado para o Fiscal';
                case eStatus.CONTROLADORIA:
                    return 'Enviado para Controladoria';
                case eStatus.CDAG:
                    return 'Enviado para CEDAG';
                case eStatus.CONTROLES_INTERNOS:
                    return 'Enviado para o Controles Internos';
                case eStatus.FINALIZADO:
                    return 'Finalizado';
                case eStatus.DEVOLVIDO:
                    return 'Devolvido para o requisitante';
                case eStatus.CLIENTE_PROSPECT:
                    return 'Criação Prospect';
                case eStatus.CANCELADO:
                    return 'Cancelamento';
                default:
                    return '-';
            }
        }

        function obterBadgeStatus(status) {
            switch(status) {
                case eStatus.PENDENTE:
                    return '<span class="badge badge-yellow">PENDENTE</span>';
                case eStatus.FISCAL:
                    return '<span class="badge badge-purple">FISCAL</span>';
                case eStatus.CONTROLADORIA:
                    return '<span class="badge badge-blue">CONTROLADORIA</span>';
                case eStatus.CDAG:
                    return '<span class="badge badge-green">CDAG</span>';
                case eStatus.CONTROLES_INTERNOS:
                    return '<span class="badge badge-pink">CONTROLES INTERNOS</span>';
                case eStatus.FINALIZADO:
                    return '<span class="badge badge-green">FINALIZADO</span>';
                case eStatus.DEVOLVIDO:
                    return '<span class="badge badge-red">DEVOLVIDO</span>';
                case eStatus.CLIENTE_PROSPECT:
                    return '<span class="badge badge-orange">CLIENTE PROPECT</span>';
                case eStatus.CANCELADO:
                    return '<span class="badge badge-gray">CANCELADO</span>';
                default:
                    return '-';
            }
        }

        function obterItemAcao(codigo, icone, texto, evento) {
            return `<a class="dropdown-item" onclick="${evento}('${codigo}')"><i class="${icone} me-1"></i> ${texto}</a>`;
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

        function adicionarAcaoHistoricoVistos(itens, codigo) {
            itens.push(
                `<hr class="dropdown-divider">`,
                obterItemAcao(codigo, 'bi bi-clock-history', 'Histórico de Vistos', 'abrirModalHistoricoVistos')
            );
        }

        function obterAcoesSolicitante(item) {
            const codigo = item.codigoOriginal;
            const status = item.status;
            const itens = [
                obterItemAcao(codigo, 'bi bi-eye', 'Visualizar', 'visualizarCadastro'),
                //obterItemAcao(codigo, 'bi bi-paperclip', 'Anexar Arquivos', 'abrirFormularioAnexo'),
            ];

            if (status == eStatus.PENDENTE || status == eStatus.DEVOLVIDO) {
                itens.splice(1, 0, obterItemAcao(codigo, 'bi bi-pencil-square', 'Editar/Unificar', 'editarCadastro'));
                itens.push(obterItemAcao(codigo, 'bi bi-ban', 'Cancelar', 'abrirModalCancelarCadastro'));
                if(item.preCadastro !== 'P') itens.push(obterItemAcao(codigo, 'bi bi-box-arrow-up-right', 'Enviar para Fiscal', 'abrirModalEnviarParaFiscal'))
            }

            adicionarAcaoHistoricoVistos(itens, codigo);
            return obterMenuAcoes(itens);
        }

        function obterAcoesCDAG(item) {
            const codigo = item.codigoOriginal;
            const status = item.status;
            const itens = [
                obterItemAcao(codigo, 'bi bi-eye', 'Visualizar', 'visualizarCadastro'),
                //obterItemAcao(codigo, 'bi bi-paperclip', 'Anexar Arquivos', 'abrirFormularioAnexo'),
            ];

            if (status == eStatus.CDAG) {
                itens.push(
                    obterItemAcao(codigo, 'bi bi-ban', 'Cancelar', 'abrirModalCancelarCadastro'),
                    obterItemAcao(codigo, 'bi bi-arrow-return-left', 'Devolver para Requisitante', 'abrirModalDevolverParaRequisitante'),
                    obterItemAcao(codigo, 'bi bi-box-arrow-up-right', 'Finalizar', 'abrirFormularioFinalizar'),
                    obterItemAcao(codigo, 'bi bi-box-arrow-up-right', 'Enviar para Controles Internos', 'abrirModalEnviarParaControlesInternos')
                );
            }

            adicionarAcaoHistoricoVistos(itens, codigo);
            return obterMenuAcoes(itens);
        }

        function obterAcoesControlesInternos(item) {
            const codigo = item.codigoOriginal;
            const status = item.status;
            const itens = [
                obterItemAcao(codigo, 'bi bi-eye', 'Visualizar', 'visualizarCadastro'),
                //obterItemAcao(codigo, 'bi bi-paperclip', 'Anexar Arquivos', 'abrirFormularioAnexo'),
            ];

            if (status == eStatus.CONTROLES_INTERNOS) {
                itens.push(
                    obterItemAcao(codigo, 'bi bi-box-arrow-up-right', 'Enviar para CDAG', 'abrirModalEnviarParaCDAG'),
                    obterItemAcao(codigo, 'bi bi-arrow-return-left', 'Devolver para Requisitante', 'abrirModalDevolverParaRequisitante')
                );
            }

            adicionarAcaoHistoricoVistos(itens, codigo);
            return obterMenuAcoes(itens);
        }
      
       function obterAcoesControladoria(item) {
            const codigo = item.codigoOriginal;
            const status = item.status;
            const itens = [
                obterItemAcao(codigo, 'bi bi-eye', 'Visualizar', 'visualizarCadastro'),
                //obterItemAcao(codigo, 'bi bi-paperclip', 'Anexar Arquivos', 'abrirFormularioAnexo'),
            ];

            if (status == eStatus.CONTROLADORIA) {
                itens.push(
                    obterItemAcao(codigo, 'bi bi-box-arrow-up-right', 'Enviar para CDAG', 'abrirModalEnviarParaCDAG'),
                    obterItemAcao(codigo, 'bi bi-arrow-return-left', 'Devolver para Requisitante', 'abrirModalDevolverParaRequisitante'),
                    obterItemAcao(codigo, 'bi bi-ban', 'Cancelar', 'abrirModalCancelarCadastro')
                );
            }

            adicionarAcaoHistoricoVistos(itens, codigo);
            return obterMenuAcoes(itens);
        }

         function obterAcoesFiscal(item) {
            const codigo = item.codigoOriginal;
            const status = item.status;
            const itens = [
                obterItemAcao(codigo, 'bi bi-eye', 'Visualizar', 'visualizarCadastro'),
                //obterItemAcao(codigo, 'bi bi-paperclip', 'Anexar Arquivos', 'abrirFormularioAnexo'),
            ];

            if (status == eStatus.FISCAL) {
                itens.push(
                    obterItemAcao(codigo, 'bi bi-box-arrow-up-right', 'Enviar para Controladoria', 'abrirFormularioEnviarParaControladoria'),
                    obterItemAcao(codigo, 'bi bi-arrow-return-left', 'Devolver para Requisitante', 'abrirModalDevolverParaRequisitante')
                );
            }

            adicionarAcaoHistoricoVistos(itens, codigo);
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

            $('#nivelUsuario').text(nivelUsuario.replace('_', ' '));

            if(nivelUsuario == 'SOLICITANTE') {
                $('#btnNovoCadastro').removeClass('d-none');
            }

            inicializarSelect2();
            carregarCadastros();
        });
    </script>
</head>
<body>
    <div class="container my-4">
        <div class="my-4">
            <h3>Cadastro Unificado</h3>
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
                        <option value="P">PENDENTE</option>
                        <option value="S">FISCAL</option>
                        <option value="C">CONTROLADORIA</option>
                        <option value="E">CDAG</option>
                        <option value="I">CONTROLES INTERNOS</option>
                        <option value="F">FINALIZADO</option>
                        <option value="R">DEVOLVIDO</option>
                        <option value="X">CANCELADO</option>
                        <option value="Z">CLIENTE PROSPECT</option>
                    </select>
                </div>

                <div class="col-12 mt-4 d-flex justify-content-between">
                    <div class="d-flex align-items-end">
                        <div class="form-check form-switch">
                            <input onchange="carregarCadastros()" checked class="form-check-input" type="checkbox" name="flagApenasPendentes" id="flagApenasPendentes" />
                            <label class="form-check-label" for="flagApenasPendentes">Apenas minhas pendências</label>
                        </div>
                    </div>

                    <div class="d-flex justify-content-end align-items-end gap-2">
                        <button onclick="limparFiltros()" type="button" class="btn btn-sm btn-outline-secondary" id="btnLimpar">
                            <i class="bi bi-x-lg me-1"></i> Limpar
                        </button>
                        <button onclick="carregarCadastros()" type="button" class="btn btn-sm btn-primary" id="btnPesquisar">
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
                <button id="btnNovoCadastro" class="d-none btn btn-sm btn-primary" onclick="novoCadastro()">
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

                <table class="table table-hover mb-0" id="gradeCadastro">
                    <thead>
                        <tr>
                            <th>#</th>
                            <th>Planta</th>
                            <th>Data Inclusão</th>
                            <th>Status</th>
                            <th>Cliente</th>
                            <th>CPF/CNPJ</th>
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

    <!-- Modal de Histórico de vistos -->
    <div class="modal fade" id="modalHistoricoVistos" tabindex="-1" aria-labelledby="modalHistoricoVistosLabel" aria-hidden="true">
        <div class="modal-dialog modal-xl">
            <div class="modal-content">
                <div class="modal-header">
                    <h5 class="modal-title" id="modalHistoricoVistosTitulo">
                        Histórico de Vistos #<span id="codigoHistoricoVistos"></span>
                    </h5>
                    <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Fechar"></button>
                </div>
                <div class="modal-body">
                    <table class="table table-hover mb-0" id="gradeHistoricoVistos">
                        <thead>
                            <tr>
                                <th>#</th>
                                <th>Data</th>
                                <th>Leadtime</th>
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
                        <tfoot>
                            <tr>
                                <td colspan="2">Total Leadtime</td>
                                <td id="totalLeadtimeVistos">0</td>
                                <td colspan="3"></td>
                            </tr>
                        </tfoot>
                    </table>
                </div>
            </div>
        </div>
    </div>
    
    <!-- Modal de confirmação de ação -->
    <div class="modal fade" id="modalAcaoCadastro" tabindex="-1" aria-labelledby="modalAcaoCadastroLabel" aria-hidden="true">
        <div class="modal-dialog modal-xl">
            <div class="modal-content">
                <div class="modal-header">
                        <h5 class="modal-title" id="modalAcaoCadastroTitulo">
                        <span id="modalAcaoCadastroTituloTexto">Confirmar ação</span> #<span id="idAcaoCadastro"></span>
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
                    <button onclick="executarAcaoCadastro()" class="btn btn-sm btn-success" id="btnEnviarResponsavel">
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