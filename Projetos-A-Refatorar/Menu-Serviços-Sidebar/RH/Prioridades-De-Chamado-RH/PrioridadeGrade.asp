<!DOCTYPE html>
<html lang="pt-br">
<head>
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

        .table-wrapper {
            overflow: auto;
            max-height: 70vh;
        }

        #gradePrioridade {
            width: max-content;
            min-width: 100%;
        }

        #gradePrioridade th,
        #gradePrioridade td {
            white-space: nowrap;
            vertical-align: top;
        }

        #gradePrioridade th:last-child,
        #gradePrioridade td:last-child {
            position: sticky;
            right: 0;
            z-index: 5;
            white-space: nowrap;
        }
        
        #gradePrioridade td:last-child {
            background: linear-gradient(to right, rgba(0,0,0,0.1) 0%, transparent 1px),#fff;
            background-origin: border-box;
            background-clip: padding-box;
        }
        
        #gradePrioridade thead th:last-child {
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
    </style>

    <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.0.2/dist/js/bootstrap.bundle.min.js" integrity="sha384-MrcW6ZMFYlzcLA8Nl+NtUVF0sA7MsXsP1UyJoMp4YLEuNSfAP+JcXn/tWtIaxVXM" crossorigin="anonymous"></script>
    <script src="https://cdn.jsdelivr.net/npm/select2@4.1.0-rc.0/dist/js/select2.min.js"></script>
    
    <script src="/corporativo/components/toast.js?v=<%=Now()%>"></script>
    <script src="/corporativo/components/confirmation.js?v=<%=Now()%>"></script>
    <script src="/corporativo/funcoes/funcoesJAVASCRIPT.js?v=<%=Now()%>"></script>
    <script src="/corporativo/components/dropdownMenu.js?v=<%=Now()%>"></script>

    <script>
        let listaPrioridades = [];

        function exibirLoading(exibir) {
            if (exibir) {
                $('#loadingOverlay').show();
            } else {
                $('#loadingOverlay').hide();
            }
        }

        function carregarPrioridades() {            
            exibirLoading(true);

            const params = {
                acao: 'listar',
            };

            $.ajax({
                url: 'PrioridadeRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: params,
                dataType: 'json',
                success: function(response) {
                    if (!response.sucesso) {
                        Toast.error(response.mensagem);
                        return;
                    }
                
                    listaPrioridades = response.dados || [];
                    renderizarGrade();
                },
                error: function(xhr, status, error) {
                    console.error('Erro ao pesquisar:', error);
                    Toast.error('Erro ao realizar a pesquisa');
                }, complete: function() {
                    exibirLoading(false);
                }
            });
        }

        function renderizarGrade() {
            const $tbody = $('#gradePrioridade tbody');
            
            $tbody.empty();
            
            if (listaPrioridades.length === 0) {
                $tbody.html(`
                    <tr>
                        <td colspan="28" class="text-center text-muted py-4">
                            Nenhum registro encontrado
                        </td>
                    </tr>
                `);
                return;
            }

            listaPrioridades.forEach(function(item) {    
                $tbody.append(`
                    <tr>
                        <td>${item.id || '-'}</td>
                        <td>${obterNomePlanta(item.codPlanta)}</td>
                        <td>${item.nivel || '-'}</td>
                        <td>${item.prazoSLA || 0} horas</td>
                        <td>${item.dataCriacao}</td>
                        <td style="text-align: center">
                            <dropdown-menu label="Ações">
                                <a class="dropdown-item" onclick="editarPrioridade('${item.id}')"><i class="bi bi-pencil me-1"></i> Editar</a>,
                                <a class="dropdown-item" onclick="confirmacaoExclusao('${item.id}')"><i class="bi bi-trash me-1"></i> Excluir</a>,
                            </dropdow-menu>
                        </td>
                    </tr>
                `);
            });
        }
   
        function novaPrioridade() {
            window.location.href = `PrioridadeFormulario.asp`;
        }

        function editarPrioridade(id) {
            window.location.href = `PrioridadeFormulario.asp?id=${id}`;
        }

        function confirmacaoExclusao(id) {
            ConfirmationModal.danger({
                message: 'Deseja realmente excluir essa prioridade?',
                type: 'danger',
                confirmText: 'Confirmar',
                cancelText: 'Cancelar',
                onConfirm: () => {
                    excluirPrioridade(id);
                },
                onCancel: () => {}
            });
        }

        function excluirPrioridade(id) {            
            exibirLoading(true);

            const params = {
                acao: 'excluir',
                id: id,
            };

            $.ajax({
                url: 'PrioridadeRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: params,
                dataType: 'json',
                success: function(response) {
                    if (!response.sucesso) {
                        Toast.error(response.mensagem);
                        return;
                    }

                    Toast.success(response.mensagem); 
                    carregarPrioridades();
                },
                error: function(xhr, status, error) {
                    console.error('Erro ao excluir:', error);
                    Toast.error('Erro ao excluir');
                }, 
                complete: function() {
                    exibirLoading(false);
                }
            });
        }

        $(document).ready(function() {
            renderizarGrade();
            carregarPrioridades();
        });
    </script>
</head>
<body>
    <div class="container my-4">
        <div class="d-flex my-4">
            <h3>Prioridades de Chamados</h3>
        </div>

        <div class="card">
            <div class="d-flex justify-content-between align-items-center p-3 border-bottom">
                <div class="d-flex gap-2">
                    <button class="btn btn-sm btn-primary" onclick="novaPrioridade()">
                        <i class="bi bi-plus-lg"></i>
                        Novo
                    </button>
                </div>
            </div>

            <div class="table-wrapper">
                <div id="loadingOverlay" class="loading-overlay" style="display: none;">
                    <div class="spinner-border text-primary" role="status">
                        <span class="visually-hidden">Carregando...</span>
                    </div>
                </div>

                <table class="table table-hover mb-0" id="gradePrioridade">
                    <thead>
                        <tr>
                            <th>#</th>
                            <th>Planta</th>
                            <th>Nível</th>
                            <th>Prazo SLA</th>
                            <th>Data de Criação</th>
                            <th style="width: 120px;" class="text-center">Ações</th>                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                    
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
</body>
</html>
