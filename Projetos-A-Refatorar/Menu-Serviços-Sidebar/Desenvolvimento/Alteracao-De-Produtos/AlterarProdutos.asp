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
        <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
        <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.0.2/dist/js/bootstrap.bundle.min.js" integrity="sha384-MrcW6ZMFYlzcLA8Nl+NtUVF0sA7MsXsP1UyJoMp4YLEuNSfAP+JcXn/tWtIaxVXM" crossorigin="anonymous"></script>
        <script src="./util/Util.js"></script>
        <script type="text/javascript">
            var produtosSelecionados = new Map(); 
            var produtosAtuais = []; 
            var timeoutBusca = null; 
            var paginaAtual = 1;
            var itensPorPagina = 25;
            
            $(document).ready(function() {
                buscarProdutos();
                
                $('form').on('submit', function(e) {
                    e.preventDefault();
                    alterarStatus();
                });
            });
            
            function buscarProdutos() {
                paginaAtual = 1;

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
                    url: 'ProdutoRepositorio.asp',
                    type: 'GET',
                    dataType: 'json',
                    data: {
                        acao: 'obterProdutos',
                        termo: termoBusca.trim().toUpperCase(),
                        pagina: paginaAtual,
                        itensPorPagina: itensPorPagina
                    },
                    success: function(response) {
                        produtosAtuais = response.produtos || [];
                        renderizarTabela();
                        atualizarPaginacao(response.totalRegistros);
                    },
                    error: function(xhr, status, error) {
                        console.error('Erro ao buscar produtos:', error);
                    }
                });
            }
            
            function renderizarTabela() {
                var tbody = $('#GradeProdutos tbody');
                tbody.empty();
                
                var selecionados = Array.from(produtosSelecionados.values());
                
                var produtosNaoSelecionados = produtosAtuais.filter(function(produto) {
                    return !produtosSelecionados.has(produto.codProduto);
                });
                
                var produtosExibir = selecionados.concat(produtosNaoSelecionados);
                
                if (produtosExibir.length === 0) {
                    $('.no-records').show();
                    return;
                }
                
                $('.no-records').hide();
                
                produtosExibir.forEach(function(produto) {
                    var produtoSelecionado = produtosSelecionados.has(produto.codProduto);
                    var produtoFixado = produtoSelecionado && !produtosAtuais.some(function(p) { return p.codProduto === produto.codProduto; });
                    var classeLinha = produtoFixado ? 'pinned-row' : '';
                    
                    var html = '<tr class="' + classeLinha + '" data-cod-produto="' + produto.codProduto + '">' +
                        '<td>' +
                            '<input type="checkbox" class="form-check-input checkbox-produto" ' +
                            'data-cod-produto="' + produto.codProduto + '" ' +
                            (produtoSelecionado ? 'checked' : '') + ' ' +
                            'onclick="toggleSelecao(\'' + produto.codProduto + '\')">' +
                        '</td>' +
                        '<td>' + produto.codCliente + ' - ' + produto.nomeCliente + '</td>' +
                        '<td>' + produto.codProduto + '</td>' +
                        '<td>' + produto.descricao + '</td>' +
                        '<td>' + obterBadgeStatus(produto.status) + '</td>' +
                    '</tr>';
                    
                    tbody.append(html);
                });
                
                atualizarContador();
                atualizarCheckboxSelectAll();
            }
            
            function toggleSelecao(codProduto) {
                var produto = produtosAtuais.find(function(p) { return p.codProduto === codProduto; });
                
                if (!produto) {
                    produto = produtosSelecionados.get(codProduto);
                }
                
                if (!produto) return;
                
                if (produtosSelecionados.has(codProduto)) {
                    produtosSelecionados.delete(codProduto);
                } else {
                    produtosSelecionados.set(codProduto, produto);
                }

                $('#pesquisa').val('');              
                realizarBusca();
            }
            
            function toggleSelectAll() {
                var checkbox = $('#checkboxSelectAll');
                var selecionar = checkbox.is(':checked');
                
                if (selecionar) {
                    produtosAtuais.forEach(function(produto) {
                        produtosSelecionados.set(produto.codProduto, produto);
                    });
                } else {
                    produtosSelecionados.clear();
                }
                
                $('#pesquisa').val('');              
                realizarBusca();
            }
            
            function atualizarCheckboxSelectAll() {
                var checkbox = $('#checkboxSelectAll');
                
                if (produtosAtuais.length === 0) {
                    checkbox.prop('checked', false);
                    checkbox.prop('indeterminate', false);
                    return;
                }
                
                var todosNaPaginaSelecionados = produtosAtuais.every(function(produto) {
                    return produtosSelecionados.has(produto.codProduto);
                });
                
                var algunsNaPaginaSelecionados = produtosAtuais.some(function(produto) {
                    return produtosSelecionados.has(produto.codProduto);
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
                var total = produtosSelecionados.size;
                var texto = total + (total === 1 ? ' item selecionado' : ' itens selecionados');
                $('#contador-selecionados').text(texto);
                
                if (total > 0) {
                    $('#card-acoes').show();
                } else {
                    $('#card-acoes').hide();
                }
            }
            
            function obterBadgeStatus(status) {
                var badges = {
                    'A': '<span class="status-badge badge-aprovado">Aprovado</span>',
                    'C': '<span class="status-badge badge-cancelado">Cancelado</span>',
                    'P': '<span class="status-badge badge-projeto">Projeto</span>',
                    'V': '<span class="status-badge badge-verificado">Verificado</span>',
                    'W': '<span class="status-badge badge-validado">Validado</span>',
                    'E': '<span class="status-badge badge-espera">Em Espera</span>'
                };
                
                return badges[status] || status;
            }
            
            function alterarStatus() {
                if (produtosSelecionados.size === 0) {
                    alert('Nenhum produto selecionado');
                    return;
                }
                
                var novoStatus = $('#status').val();
                var produtosIds = Array.from(produtosSelecionados.keys());
                                
                if (!confirm('Deseja realmente alterar o status de ' + produtosSelecionados.size + ' produto(s)?')) {
                    return;
                }

                
                $.ajax({
                    url: 'ProdutoRepositorio.asp',
                    type: 'POST',
                    dataType: 'json',
                    data: {
                        acao: 'alterarStatus',
                        produtosIds: (produtosIds.map(function(codProduto) { return "'" + codProduto.trim() + "'"; }).join(',')),
                        status: novoStatus,
                    },
                    success: function(response) {
                        if (response.sucesso) {
                            alert('Status alterado com sucesso!');
                            produtosSelecionados.clear();
                            realizarBusca();
                        } else {
                            alert(response.mensagem || 'Erro ao alterar status');
                        }
                    },
                    error: function(xhr, status, error) {
                        console.error('Erro ao alterar status:', error);
                        alert('Erro ao alterar status dos produtos');
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
            body{
                background-color: #f5f5f5;
                font-size: .9rem;
            }
            .card {
                padding: 0rem;
            }
            .form-label {
                font-weight: 600;
            }
            .table-wrapper {
                overflow: auto;
                max-height: 70vh;
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
                padding: 40px 10px;
                text-align: center;
                background: white;
                color: #6c757d;
                display: none;
            }
            .status-badge {
                display: inline-block;
                text-align: center;
                padding: 6px 12px;
                border-radius: 6px;
                font-size: 0.75rem;
                font-weight: 600;
            }
            .badge-aprovado {
                background: #d4edda;
                color: #155724;
            }
            .badge-cancelado {
                background: #f8d7da;
                color: #721c24;
            }
            .badge-projeto {
                background: #d1ecf1;
                color: #0c5460;
            }
            .badge-verificado {
                background: #fff3cd;
                color: #856404;
            }
            .badge-validado {
                background: #d1d3e2;
                color: #383d41;
            }
            .badge-espera {
                background: #e2e3e5;
                color: #383d41;
            }
            .pinned-row {
                background-color: rgba(13, 110, 253, 0.05);
                border-left: 4px solid #0d6efd;
            }
            #card-acoes {
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
            <h3 class="mt-4 mb-4">Alterar Produtos</h3>

            <div class="col-12 mb-4">
                <label for="pesquisa" class="form-label">Busque por código ou descrição</label>
                <input type="text" class="form-control form-control-sm" id="pesquisa" name="pesquisa" placeholder="Buscar produtos..." oninput="buscarProdutos()">
            </div>

            <div class="card p-3 mb-4" id="card-acoes">
                <form>
                    <div class="row g-3">
                        <div class="col-md-4">
                            <div class="d-flex align-items-center h-100">
                                <span id="contador-selecionados">0 itens selecionados</span>
                            </div>
                        </div>
                       
                        <div class="col-md-4">
                            <select class="form-select form-select-sm" name="status" id="status" required>
                                <option selected value="A">Aprovado - A</option>
                                <option value="C">Cancelado - C</option>
                                <option value="P">Projeto - P</option>
                                <option value="V">Verificado - V</option>
                                <option value="W">Validado - W</option>
                                <option value="E">Em Espera - E</option>
                            </select>
                        </div>

                        <div class="col-md-4"> 
                            <button type="submit" class="btn btn-sm btn-primary form-control">
                                Atualizar Status
                            </button>
                        </div>
                    </div>
                </form>
            </div>
            
            <div class="card">
                <div class="table-wrapper">
                    <table class="table table-hover mb-0" name="GradeProdutos" id="GradeProdutos">
                        <thead>
                            <tr>
                                <th style="width: 50px;">
                                    <input type="checkbox" class="form-check-input" id="checkboxSelectAll" onclick="toggleSelectAll()">
                                </th>
                                <th>Cliente</th>
                                <th>Código</th>
                                <th>Descrição</th>
                                <th>Status</th>
                            </tr>              
                        </thead>
                        <tbody>
                            <!-- Inserido dinamicamente -->
                        </tbody>
                    </table>
                    <div class="no-records">
                        <p class="mb-0">Nenhum registro encontrado</p>
                    </div>
                </div>
            </div>
            
            <div class="d-flex justify-content-between align-items-center mt-3">
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
    </body>
</html>