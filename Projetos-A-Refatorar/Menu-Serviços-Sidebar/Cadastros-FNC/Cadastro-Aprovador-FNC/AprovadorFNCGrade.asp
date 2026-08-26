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
            max-width: 90%;
        }

        .table-wrapper {
            overflow: auto;
            max-height: 70vh;
        }

        .truncate-text {
            max-width: 250px;
            white-space: nowrap;
            overflow: hidden;
            text-overflow: ellipsis;
            cursor: help;
        }

        #gradeAprovadores {
            width: max-content;
            min-width: 100%;
        }

        #gradeAprovadores th,
        #gradeAprovadores td {
            white-space: nowrap;
            vertical-align: top;
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

    <script>
        let idFnc = 0;
        let dadosAprovadores = [];
        let carregando = false;

        function exibirLoading(exibir) {
            carregando = exibir;
            if (exibir) {
                $('#loadingOverlay').show();
            } else {
                $('#loadingOverlay').hide();
            }
        }

        function carregarAprovadores() {            
            exibirLoading(true);

            const params = {
                acao: 'listar',
                idFnc: idFnc
            };

            $.ajax({
                url: 'AprovadorFNCRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: params,
                dataType: 'json',
                success: function(response) {
                    exibirLoading(false);

                    if (!response.sucesso) {
                        Toast.error(response.mensagem);
                        return;
                    }
                
                    dadosAprovadores = response.dados || [];
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
            const $tbody = $('#gradeAprovadores tbody');
            
            $tbody.empty();
            
            if (dadosAprovadores.length === 0) {
                $tbody.html(`
                    <tr>
                        <td colspan="28" class="text-center text-muted py-4">
                            Nenhum registro encontrado
                        </td>
                    </tr>
                `);
                return;
            }

            dadosAprovadores.forEach(function(item) {         
                $tbody.append(`
                    <tr>
                        <td>${item.idUsuario || '-'}</td>    
                        <td>${item.nome || '-'}</td>    
                        <td>${item.email || '-'}</td>    
                        <td>${item.dataInclusao || '-'}</td>
                        <td style="text-align:center">
                            <button type="button" class="btn btn-sm btn-primary" onclick="navegarParaConfiguracoes(${item.idUsuario})" title="Configurar">
                                <i class="bi bi-gear"></i>
                            </button>
                            <button type="button" class="btn btn-sm btn-danger" onclick="excluirAprovador(${item.idUsuario})" title="Excluir">
                                <i class="bi bi-trash"></i>
                            </button>
                        </td>
                    </tr>
                `);
            });
        }
   
        function novoAprovador() {
            window.location.href = `AprovadorFNCFormulario.asp?`;
        }

        function navegarParaConfiguracoes(idUsuario) {
            window.location.href = `../AprovadorParametro/AprovadorFNCParametroGrade.asp?idUsuario=${idUsuario}`;
        }

        function excluirAprovador(idUsuario) {
            ConfirmationModal.danger({
                message: 'Deseja realmente excluir este aprovador?',
                type: 'danger',
                confirmText: 'Confirmar',
                cancelText: 'Cancelar',
                onConfirm: () => {
                    confirmarExclusao(idUsuario);
                },
                onCancel: () => {}
            });
        }

        function confirmarExclusao(idUsuario) {            
            exibirLoading(true);

            const params = {
                acao: 'excluir',
                idUsuario: idUsuario,
            };

            $.ajax({
                url: 'AprovadorFNCRepositorio.asp?v=<%=Now()%>',
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
                    carregarAprovadores();
                },
                error: function(xhr, status, error) {
                    exibirLoading(false);
                    console.error('Erro ao excluir:', error);
                    Toast.error('Erro ao excluir');
                }
            });
        }

        $(document).ready(function() {
            renderizarGrade();
            carregarAprovadores();
        });
    </script>
</head>
<body>
    <div class="container my-4">
        <div class="d-flex my-4">
            <h3>Aprovadores de FNC</h3>
        </div>

        <div class="card">
            <div class="d-flex justify-content-between align-items-center p-3 border-bottom">
                <div class="d-flex gap-2">
                    <button class="btn btn-sm btn-primary" onclick="novoAprovador()">
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

                <table class="table table-hover mb-0" id="gradeAprovadores">
                    <thead>
                        <tr>
                            <th>#</th>
                            <th>Nome</th>
                            <th>Email</th>
                            <th>Data Inclusão</th>    
                            <th style="width: 100px;">Ações</th>                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                    
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
