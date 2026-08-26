<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file = "../cnn2/cnn2.asp" -->
<!--#include File = "../funcoes/funcoes.asp"-->
<!--#include file="../includes/variaveis.asp" -->
<%
Response.Charset = "UTF-8"
Response.CodePage = 65001
Response.Buffer  = true
Response.Expires = 0
Session.LCID = 1046

Dim solicitacoes()

'Variaveis auxiliares'
Dim idx
idx = 0
totalSolicitacoes = 0

Dim idUsuarioLogado
idUsuarioLogado = Session("CODIGO")

Dim usuarioLogado
Set usuarioLogado = CreateObject("Scripting.Dictionary")
%>

<!DOCTYPE html>
<html lang="pt-br">
    <head>
        <meta charset="utf-8"/>
        <meta name="viewport" content="width=device-width, initial-scale=1"/>
        <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.0.2/dist/css/bootstrap.min.css" rel="stylesheet" integrity="sha384-EVSTQN3/azprG1Anm3QDgpJLIm9Nao0Yz1ztcQTwFspd3yD65VohhpuuCOmLASjC" crossorigin="anonymous" />
        <link rel="stylesheet" href="https://use.fontawesome.com/releases/v5.7.0/css/all.css" integrity="sha384-lZN37f5QGtY3VHgisS14W3ExzMWZxybE1SJSEsQp9S+oqd12jhcu+A56Ebc1zFSJ" crossorigin="anonymous" />    
        <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.1/font/bootstrap-icons.css" />

        <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
        <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.0.2/dist/js/bootstrap.bundle.min.js" integrity="sha384-MrcW6ZMFYlzcLA8Nl+NtUVF0sA7MsXsP1UyJoMp4YLEuNSfAP+JcXn/tWtIaxVXM" crossorigin="anonymous"></script>
         
        <script src="/corporativo/components/toast.js?v=<%=Now()%>"></script>
        <script src="/corporativo/components/confirmation.js?v=<%=Now()%>"></script>
        <script src="/corporativo/funcoes/funcoesJAVASCRIPT.js?v=<%=Now()%>"></script>
        
        <script type="text/javascript">
            var orcamentosSelecionados = new Map(); 
            var orcamentosAtuais = []; 
            var timeoutBusca = null; 
            var paginaAtual = 1;
            var itensPorPagina = 25;
            
            $(document).ready(function() {
                buscarOrcamentos();
            });

            function navegarParaGrade() {
                window.location.href = `OrcamentoGrade.asp`;
            }
            
            function buscarOrcamentos() {
                if (timeoutBusca) {
                    clearTimeout(timeoutBusca);
                }
                
                timeoutBusca = setTimeout(function() {
                    realizarBusca();
                }, 500);
            }
            
            function realizarBusca() {
                var termoBusca = $('#pesquisa').val();
                
                $.ajax({
                    url: 'OrcamentoDesativarRepositorio.asp',
                    type: 'GET',
                    dataType: 'json',
                    data: {
                        acao: 'listarOrcamentos',
                        termo: termoBusca.trim().toUpperCase(),
                        pagina: paginaAtual,
                        itensPorPagina: itensPorPagina
                    },
                    success: function(response) {
                        orcamentosAtuais = response.dados || [];
                        renderizarTabela();
                        atualizarPaginacao(response.totalRegistros);
                    },
                    error: function(xhr, status, error) {
                        console.error('Erro ao buscar orçamentos:', error);
                    }
                });
            }
            
            function renderizarTabela() {
                var tbody = $('#gradeOrcamento tbody');
                tbody.empty();
                
                var selecionados = Array.from(orcamentosSelecionados.values());
                
                var orcamentosNaoSelecionados = orcamentosAtuais.filter(function(orcamento) {
                    return !orcamentosSelecionados.has(orcamento.idOrcamento);
                });
                
                var orcamentosExibir = selecionados.concat(orcamentosNaoSelecionados);
                
                if (orcamentosExibir.length === 0) {
                    $('.no-records').show();
                    return;
                }
                
                $('.no-records').hide();
                
                orcamentosExibir.forEach(function(orcamento) {
                    var orcamentoSelecionado = orcamentosSelecionados.has(orcamento.idOrcamento);
                    var classeLinha = orcamentoSelecionado ? 'pinned-row' : '';

                    var html = '<tr class="' + classeLinha + '" data-id-orcamento="' + orcamento.idOrcamento + '">' +
                        '<td>' +
                            '<input type="checkbox" class="form-check-input checkbox-orcamento" ' +
                            'data-id-orcamento="' + orcamento.idOrcamento + '" ' +
                            (orcamentoSelecionado ? 'checked' : '') + ' ' +
                            'onclick="toggleSelecao(\'' + orcamento.idOrcamento + '\')">' +
                        '</td>' +
                        '<td>' + orcamento.idOrcamento + '</td>' +
                        '<td>' + orcamento.numeroOrcamento + '</td>' +
                        '<td>' + orcamento.dataInclusao + '</td>' +
                        '<td>' + orcamento.dataValidadeDe + ' - ' + orcamento.dataValidadeAte + '</td>' +
                        '<td>' + orcamento.planta + '</td>' +
                        '<td>' + orcamento.nomeCliente + '</td>' +
                        '<td>' + orcamento.cnpj + '</td>' +
                        '<td>' + orcamento.cidade + '</td>' +
                        '<td>' + orcamento.representante + '</td>' +
                    '</tr>';
                    
                    tbody.append(html);
                });
                
                atualizarContador();
                atualizarCheckboxSelectAll();
            }
            
            function toggleSelecao(idOrcamento) {
                var orcamento = orcamentosAtuais.find(function(p) { return p.idOrcamento === idOrcamento; });
                
                if (!orcamento) {
                    orcamento = orcamentosSelecionados.get(idOrcamento);
                }
                
                if (!orcamento) return;
                
                if (orcamentosSelecionados.has(idOrcamento)) {
                    orcamentosSelecionados.delete(idOrcamento);
                } else {
                    orcamentosSelecionados.set(idOrcamento, orcamento);
                }

                realizarBusca();
            }
            
            function toggleSelectAll() {
                var checkbox = $('#checkboxSelectAll');
                var selecionar = checkbox.is(':checked');
                
                if (selecionar) {
                    orcamentosAtuais.forEach(function(orcamento) {
                        orcamentosSelecionados.set(orcamento.idOrcamento, orcamento);
                    });
                } else {
                    orcamentosSelecionados.clear();
                }
                
                realizarBusca();
            }
            
            function atualizarCheckboxSelectAll() {
                var checkbox = $('#checkboxSelectAll');
                
                if (orcamentosAtuais.length === 0) {
                    checkbox.prop('checked', false);
                    checkbox.prop('indeterminate', false);
                    return;
                }
                
                var todosNaPaginaSelecionados = orcamentosAtuais.every(function(orcamento) {
                    return orcamentosSelecionados.has(orcamento.idOrcamento);
                });
                
                var algunsNaPaginaSelecionados = orcamentosAtuais.some(function(orcamento) {
                    return orcamentosSelecionados.has(orcamento.idOrcamento);
                });
                
                if (todosNaPaginaSelecionados) {
                    checkbox.prop('checked', true);
                    checkbox.prop('indeterminate', false);
                } else if (algunsNaPaginaSelecionados) {
                    checkbox.prop('checked', false);
                    checkbox.prop('indeterminate', true);
                } else {
                    checkbox.prop('checked', false);
                    checkbox.prop('indeterminate', false);
                }
            }
            
            function atualizarContador() {
                var total = orcamentosSelecionados.size;
                var texto = total + (total === 1 ? ' item selecionado' : ' itens selecionados');
                $('#contadorSelecionados').text(texto);
                
                if (total > 0) {
                    $('#divAcoes').show();
                } else {
                    $('#divAcoes').hide();
                }
            }

            function desativarOrcamentos() {
                if (orcamentosSelecionados.size === 0) {
                    Toast.warning('Nenhum orcamento selecionado');
                    return;
                }
                                
                ConfirmationModal.danger({
                    message: 'Deseja realmente desativar ' + orcamentosSelecionados.size + ' orcamento(s)?',
                    confirmText: 'Confirmar',
                    cancelText: 'Cancelar',
                    onConfirm: () => {
                        confirmarDesativarOrcamentos();
                    },
                    onCancel: () => {}
                });
            }
            
            function confirmarDesativarOrcamentos() {
                var idsOrcamento = Array.from(orcamentosSelecionados.keys());

                $.ajax({
                    url: 'OrcamentoDesativarRepositorio.asp',
                    type: 'POST',
                    dataType: 'json',
                    data: {
                        acao: 'desativarOrcamentos',
                        idsOrcamento: (idsOrcamento.map(function(idOrcamento) { return "'" + idOrcamento.trim() + "'"; }).join(',')),
                    },
                    success: function(response) {
                        if (!response.sucesso) {
                            Toast.error(response.mensagem || 'Erro ao desativar orçamentos');
                        }

                        Toast.success('Orçamentos desativados com sucesso!');
                        orcamentosSelecionados.clear();
                        realizarBusca();
                    },
                    error: function(xhr, status, error) {
                        console.error('Erro ao desativar orçamentos:', error);
                        Toast.error('Erro ao carregar orçamentos');
                    }
                });
            }
            
            function atualizarPaginacao(totalRegistros) {
                var totalPaginas = Math.ceil(totalRegistros / itensPorPagina);
                $('#info-paginacao').text('Página ' + paginaAtual + ' de ' + totalPaginas);
                
                $('#btn-anterior').prop('disabled', paginaAtual === 1);
                $('#btn-proxima').prop('disabled', paginaAtual === totalPaginas || totalPaginas === 0);
            }
            
            function paginaAnterior() {
                if (paginaAtual > 1) {
                    paginaAtual--;
                    realizarBusca();
                }
            }
            
            function proximaPagina() {
                paginaAtual++;
                realizarBusca();
            }
        </script>
        
        <style type="text/css">
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

            #gradeOrcamento th,
            #gradeOrcamento td {
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
            .no-records {
                text-align: center;
                color: #6c757d;
                display: none;
            }
            .pinned-row {
                background-color: rgba(13, 110, 253, 0.05) !important;
                border-left: 4px solid #0d6efd !important;
            }
            #divAcoes {
                display: none;
            }
            .form-check-input {
                width: 1.2rem;
                height: 1.2rem;
                cursor: pointer;
            }
        </style>
    </head>
    <body>
        <div class="container">
            <div class="d-flex my-4">
                <button onclick="navegarParaGrade()" class="btn btn-sm btn-light me-2 fs-5"> 
                    <i class="bi bi-arrow-left"></i> 
                </button>
                <h3 class="mb-0">Desativar Orçamentos</h3>
            </div>

            <div class="mb-4">
                <label for="pesquisa" class="form-label">Busque por Id, número do Orçamento, FT ou Cliente</label>
                <input type="text" class="form-control form-control-sm" id="pesquisa" name="pesquisa" placeholder="Buscar orçamentos..." oninput="buscarOrcamentos()">
            </div>

            <div class="card">
                <div class="d-flex p-3 gap-2 align-items-center">
                    <button onclick="desativarOrcamentos()" type="button" class="btn btn-danger btn-sm" id="btnDesativar">
                        <i class="bi bi-ban me-1"></i> Desativar Registros
                    </button>

                    <span class="badge bg-primary"><span id="contadorSelecionados">0</span></span>
                </div>

                <div class="table-wrapper">
                    <div id="loadingOverlay" class="loading-overlay" style="display: none;">
                        <div class="spinner-border text-primary" role="status">
                            <span class="visually-hidden">Carregando...</span>
                        </div>
                    </div>

                    <table class="table table-hover mb-0" name="gradeOrcamento" id="gradeOrcamento">
                       <thead>
                            <tr>
                                <th style="width: 50px;">
                                    <input type="checkbox" class="form-check-input" id="checkboxSelectAll" onclick="toggleSelectAll()">
                                </th>
                                <th>ID</th>
                                <th>Nº Orçamento</th>
                                <th>Data de Emissão</th>
                                <th>Data de Vigência</th>
                                <th>Planta</th>
                                <th>Cliente</th>
                                <th>CNPJ</th>
                                <th>Cidade/UF</th>
                                <th>Vendedor</th>
                            </tr>              
                        </thead>
                        <tbody>
                            <!-- Inserido dinamicamente -->
                        </tbody>

                        <tfoot>
                            
                        </tfoot>
                    </table>

                    <div class="no-records">
                        <p class="mb-0 py-4">Nenhum registro encontrado</p>
                    </div>

                    <div class="d-flex justify-content-between align-items-center p-2 border-top">
                        <div>
                            <small class="text-muted" id="info-paginacao">Página 1 de 1</small>
                        </div>
                        <div>
                            <button class="btn btn-sm btn-outline-secondary" id="btn-anterior" onclick="paginaAnterior()">
                                <i class="fas fa-chevron-left"></i> Anterior
                            </button>
                            <button class="btn btn-sm btn-outline-secondary ms-2" id="btn-proxima" onclick="proximaPagina()">
                                Próxima <i class="fas fa-chevron-right"></i>
                            </button>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    </body>
</html>