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

        .form-label {
            font-weight: 600;
        }

        .table-wrapper {
            overflow: auto;
            max-height: 70vh;
        }

         .icon-wrapper {
            width: 48px;
            height: 48px;
            border-radius: 8px;
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 22px;
        }
        
        .icon-wrapper.purple {
            background-color: #ede9fe;
            color: #8b5cf6;
        }

        .icon-wrapper.red {
            background-color: #fcdbdb;
            color: #fa4a4a;
        }

        .icon-wrapper.orange {
            background-color: #fce8db;
            color: #ff7b00;
        }
        
        .icon-wrapper.yellow {
            background-color: #fef9c3;
            color: #ca8a04;
        }
        
        .icon-wrapper.green {
            background-color: #d1fae5;
            color: #10b981;
        }
        
        .icon-wrapper.blue {
            background-color: #dbeafe;
            color: #3b82f6;
        }

        #gradeChamado {
            width: max-content;
            min-width: 100%;
        }

        #gradeChamado th,
        #gradeChamado td {
            white-space: nowrap;
            vertical-align: top;
        }

        #gradeChamado th:last-child,
        #gradeChamado td:last-child {
            position: sticky;
            right: 0;
            z-index: 5;
            white-space: nowrap;
        }
        
        #gradeChamado td:last-child {
            background: linear-gradient(to right, rgba(0,0,0,0.1) 0%, transparent 1px),#fff;
            background-origin: border-box;
            background-clip: padding-box;
        }
        
        #gradeChamado thead th:last-child {
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
    </style>

    <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.0.2/dist/js/bootstrap.bundle.min.js" integrity="sha384-MrcW6ZMFYlzcLA8Nl+NtUVF0sA7MsXsP1UyJoMp4YLEuNSfAP+JcXn/tWtIaxVXM" crossorigin="anonymous"></script>
    <script src="https://cdn.jsdelivr.net/npm/select2@4.1.0-rc.0/dist/js/select2.min.js"></script>

    <script src="/corporativo/components/toast.js?v=<%=Now()%>"></script>
    <script src="/corporativo/components/confirmation.js?v=<%=Now()%>"></script>
    <script src="/corporativo/funcoes/funcoesJAVASCRIPT.js?v=<%=Now()%>"></script>
    <script src="/corporativo/components/dropdownMenu.js?v=<%=Now()%>"></script>
    <script src="../Compartilhado/funcoes.js?v=<%=Now()%>"></script>

    <script>
        let listaChamados = [];
        let dadosKPI = {}
        
        function carregarDadosKPI() {              
            const params = {
                acao: 'obterDadosKPI',
            };

            $.ajax({
                url: 'ChamadoRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: params,
                dataType: 'json',
                success: function(response) {
                    if (!response.sucesso) {
                        Toast.error(response.mensagem);
                        return;
                    }

                    dadosKPI = response.dados;
                    renderizarCartoesKPI();
                },
                error: function(xhr, status, error) {
                    console.error('Erro ao pesquisar:', error);
                    Toast.error('Erro ao realizar a pesquisa');
                }
            });
        }

        function carregarChamados() {   
            const params = {
                acao: 'listarFila',
            };

            $.ajax({
                url: 'ChamadoRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: params,
                dataType: 'json',
                success: function(response) {
                    if (!response.sucesso) {
                        Toast.error(response.mensagem);
                        return;
                    }
                
                    listaChamados = response.dados || [];
                    renderizarGrade();
                },
                error: function(xhr, status, error) {
                    console.error('Erro ao pesquisar:', error);
                    Toast.error('Erro ao realizar a pesquisa');
                }
            });
        }

        function renderizarCartoesKPI() {
            if(!dadosKPI) {
                $('#kpiEmAndamento').text('0');
                $('#kpiResolvidosHoje').text('0');
                $('#kpiSLAVencendo').text('0');
                $('#kpiSLAVencido').text('0');
                return;
            }

            $('#kpiEmAndamento').text(dadosKPI.emAndamento);
            $('#kpiResolvidosHoje').text(dadosKPI.resolvidosHoje);
            $('#kpiSLAVencendo').text(dadosKPI.slaVencendo);
            $('#kpiSLAVencido').text(dadosKPI.slaVencido);
        }

        function renderizarGrade() {
            const $tbody = $('#gradeChamado tbody');
            $tbody.empty();

            if (listaChamados.length === 0) {
                $tbody.html(`
                    <tr>
                        <td colspan="28" class="text-center text-muted py-4">
                            Nenhum registro encontrado
                        </td>
                    </tr>
                `);
                return;
            }
            

            listaChamados.forEach(function(item) {    
                $tbody.append(`
                    <tr>
                        <td>${item.id || '-'}</td>
                        <td>${item.titulo || '-'}</td>
                        <td>${item.nomeSolicitante}</td>
                        <td>${obterBadgePrioridade(item.nivelPrioridade)}</td>
                        <td>${obterBadgeStatus(item.status)}</td>
                        <td>${item.nomeAtendente || 'Nenhum atendente'}</td>
                        <td>${item.dataCriacao}</td>
                        <td>${item.dataPrevista}</td>
                        <td style="text-align: center">
                            <dropdown-menu label="Ações">
                                <a class="dropdown-item" onclick="visualizarChamado('${item.id}')"><i class="bi bi-eye me-1"></i> Visualizar</a>,
                            </dropdow-menu>
                        </td>
                    </tr>
                `);
            });
        }

        function formatarHoras(min) {
            const horas = Math.floor(min / 60);
            const minutos = min % 60;

            if (horas === 0) return `${minutos}min`;
            if (minutos === 0) return `${horas}h`;

            return `${horas}h ${minutos}min`;
        }

        function visualizarChamado(id) {
            window.location.href = `ChamadoDetalhes.asp?id=${id}`;
        }


        function inicializarSelect2() {
            $('.select2').select2({ theme: 'bootstrap-5', width: '100%', dropdownCssClass: "z-index-9999", allowClear: true });
        }

        $(document).ready(function() {
            inicializarSelect2();

            renderizarGrade();
            carregarDadosKPI();
            carregarChamados();

            setInterval(() => { 
                carregarDadosKPI();
                carregarChamados();
            }, 30000);
        });
    </script>
</head>
<body>
    <div class="container my-4">
        <div class="d-flex my-4">
            <h3>Fila de Chamados</h3>
        </div>

        <div class="row g-3 mb-4">
            <div class="col-6 col-lg-3">
                <div class="card shadow-sm border-0 h-100 border-start">
                    <div class="card-body d-flex justify-content-between align-items-center">
                        <div>
                            <div class="text-secondary mb-1">Em andamento</div>
                            <h2 id="kpiEmAndamento">0</h2>
                        </div>
                        <div class="icon-wrapper yellow">
                            <i class="bi bi-clock"></i>
                        </div>
                    </div>
                </div>
            </div>
            <div class="col-6 col-lg-3">
                <div class="card shadow-sm border-0 h-100 border-start">
                    <div class="card-body d-flex justify-content-between align-items-center">
                        <div>
                            <div class="text-secondary mb-1">Resolvidos Hoje</div>
                            <h2 id="kpiResolvidosHoje">0</h2>
                        </div>
                        <div class="icon-wrapper green">
                            <i class="bi bi-check-circle"></i>
                        </div>
                    </div>
                </div>
            </div>
            <div class="col-6 col-lg-3">
                <div class="card shadow-sm border-0 h-100 border-start">
                    <div class="card-body d-flex justify-content-between align-items-center">
                        <div>
                            <div class="text-secondary mb-1">SLA vencendo (menos de 4h restantes)</div>
                            <h2 id="kpiSLAVencendo">0</h2>
                        </div>
                        <div class="icon-wrapper orange">
                            <i class="bi bi-exclamation-circle"></i>
                        </div>
                    </div>
                </div>
            </div>
            <div class="col-6 col-lg-3">
                <div class="card shadow-sm border-0 h-100 border-start">
                    <div class="card-body d-flex justify-content-between align-items-center">
                        <div>
                            <div class="text-secondary mb-1">SLA Vencido</div>
                            <h2 id="kpiSLAVencido">0</h2>
                        </div>
                        <div class="icon-wrapper red">
                            <i class="bi bi-exclamation-triangle"></i>
                        </div>
                    </div>
                </div>
            </div>
        </div>
       
        <div class="card">
            <div class="table-wrapper">
                <table class="table table-hover mb-0" id="gradeChamado">
                    <thead>
                        <tr>
                            <th>#</th>
                            <th>Assunto</th>
                            <th>Solicitante</th>
                            <th>Prioridade</th>
                            <th>Status</th>
                            <th>Atendente</th>
                            <th>Data Criação</th>
                            <th>Data Prevista</th>
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
