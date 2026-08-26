<!DOCTYPE html>
<html lang="pt-br">
<head>
    <title>Smurfit Kappa - Dashboard de Chamados</title>
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

        .chart-container {
            position: relative;
            height: 200px;
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

        /* Gráfico de porcentagem custom */
        .chart-wrap { position: relative; width: 180px; height: 180px; }
        .chart-center { position: absolute; inset: 0; pointer-events: none; }
        .bar-fill { height: 10px; background: rgba(25, 135, 84, 1); border-radius: 3px; }
        .bar-bg { height: 10px; background: #efefef; overflow: hidden; }

        thead {
            vertical-align: middle !important;
            position: sticky; 
            inset-block-start: 0; 
            background: #f5f5f5;
            z-index: 10;
        }

        td, th {
            padding: 8px 12px !important;
            white-space: nowrap;
            vertical-align: top;
        }

        th {
            font-weight: 600;
            color: #495057;
        }

        tr:hover {
            background: #f5f5f5;
        }
    </style>

    <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.0.2/dist/js/bootstrap.bundle.min.js" integrity="sha384-MrcW6ZMFYlzcLA8Nl+NtUVF0sA7MsXsP1UyJoMp4YLEuNSfAP+JcXn/tWtIaxVXM" crossorigin="anonymous"></script>
    <script src="https://cdn.jsdelivr.net/npm/select2@4.1.0-rc.0/dist/js/select2.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.1/dist/chart.umd.min.js"></script>
    <script src="/corporativo/components/toast.js?v=<%=Now()%>"></script>
    <script src="/corporativo/funcoes/funcoesJAVASCRIPT.js?v=<%=Now()%>"></script>

    <script>
        let dadosDashboard = null;
        let carregando = false;
        let tipoGraficoCategoria = 'tipo';
        let charts = {};

        const coresGrafico = [
            'rgba(13, 110, 253, 0.55)',      
            'rgba(25, 135, 84, 0.55)',       
            'rgba(253, 126, 20, 0.55)',      
            'rgba(220, 53, 69, 0.55)',       
            'rgba(111, 66, 193, 0.55)',     
            'rgba(32, 201, 151, 0.55)',      
            'rgba(13, 202, 240, 0.55)',     
            'rgba(255, 193, 7, 0.55)',      
            'rgba(226, 7, 255, 0.55)',      
            'rgba(214, 51, 132, 0.55)',     
            'rgba(73, 80, 87, 0.55)',        
            'rgba(173, 181, 189, 0.55)',     
            'rgba(255, 87, 51, 0.55)',      
            'rgba(156, 39, 176, 0.55)',     
            'rgba(0, 150, 136, 0.55)',      
            'rgba(63, 81, 181, 0.55)',    
            'rgba(255, 152, 0, 0.55)',      
            'rgba(121, 85, 72, 0.55)',     
            'rgba(96, 125, 139, 0.55)',     
            'rgba(233, 30, 99, 0.55)',   
            'rgba(76, 175, 80, 0.55)',   
            'rgba(255, 235, 59, 0.55)',    
            'rgba(103, 58, 183, 0.55)',    
            'rgba(0, 188, 212, 0.55)'      
        ];

        const bordasGrafico = [
            'rgba(10, 87, 202, 1.0)',      
            'rgba(20, 108, 67, 1.0)',      
            'rgba(202, 101, 16, 1.0)',      
            'rgba(176, 42, 55, 1.0)',      
            'rgba(89, 53, 154, 1.0)',      
            'rgba(26, 161, 121, 1.0)',      
            'rgba(10, 162, 192, 1.0)',      
            'rgba(204, 154, 6, 1.0)',      
            'rgba(226, 7, 255, 1.0)',      
            'rgba(171, 41, 106, 1.0)',      
            'rgba(58, 64, 70, 1.0)',        
            'rgba(138, 145, 151, 1.0)',    
            'rgba(204, 70, 41, 1.0)',      
            'rgba(125, 31, 141, 1.0)',      
            'rgba(0, 120, 109, 1.0)',      
            'rgba(50, 65, 145, 1.0)',       
            'rgba(204, 122, 0, 1.0)',      
            'rgba(97, 68, 58, 1.0)',      
            'rgba(77, 100, 111, 1.0)',    
            'rgba(186, 24, 79, 1.0)',     
            'rgba(61, 140, 64, 1.0)',    
            'rgba(204, 188, 47, 1.0)',     
            'rgba(82, 46, 146, 1.0)',     
            'rgba(0, 150, 170, 1.0)'     
        ];

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

        function campoPlantaAlterado() {
            const codPlanta = $('#filtroFilial').val();
            carregarDepartamentos(codPlanta);
            carregarAtendentes(codPlanta);
        }

        function campoDepartamentoAlterado() {
            const codPlanta = $('#filtroFilial').val();
            const idDepartamento = $('#filtroDepartamento').val();
            carregarCategorias(codPlanta, idDepartamento);
        }

        function carregarDepartamentos(codPlanta) {
            const $select = $('#filtroDepartamento');
            $select.empty().append('<option value="">Todos</option>');

            if(!codPlanta) {
                $select.prop('disabled', true);
                return;
            }

            $select.prop('disabled', false);

            const params = {
                acao: 'listar',
                codPlanta: codPlanta,
            };

            return $.ajax({
                url: '../Departamento/DepartamentoRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: params,
                dataType: 'json',
                success: function(response) {
                    if (!response.sucesso) {
                        console.error("Erro ao carregar departamentos: " + response.mensagem);
                        return;
                    }

                    response.dados.forEach(function(item) {
                        $select.append($('<option>', { value: item.id, text: item.nome }));
                    });
                },
                error: function(xhr, status, error) {
                    console.error('Erro ao carregar departamentos:', error);
                }
            });
        }

        function carregarCategorias(codPlanta, idDepartamento) {
            const $select = $('#filtroCategoria');
            $select.empty().append('<option value="">Todas</option>');

            if(!codPlanta || !idDepartamento) {
                $select.prop('disabled', true);
                return;
            }

            $select.prop('disabled', false);

            const params = {
                acao: 'listar',
                codPlanta: codPlanta,
                idDepartamento: idDepartamento,
            };

            return $.ajax({
                url: '../Categoria/CategoriaRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: params,
                dataType: 'json',
                success: function(response) {
                    if (!response.sucesso) {
                        console.error("Erro ao carregar categorias: " + response.mensagem);
                        return;
                    }

                    response.dados.forEach(function(item) {
                        $select.append($('<option>', { value: item.id, text: item.nome }));
                    });
                },
                error: function(xhr, status, error) {
                    console.error('Erro ao carregar categorias:', error);
                }
            });
        }

        function carregarAtendentes(codPlanta) {
            const $select = $('#filtroAtendente');
            $select.empty().append('<option value="">Todos</option>');

            if(!codPlanta) {
                $select.prop('disabled', true);
                return;
            }

            $select.prop('disabled', false);

            const params = {
                acao: 'listar',
                codPlanta: codPlanta,
            };

            return $.ajax({
                url: '../Atendente/AtendenteRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: params,
                dataType: 'json',
                success: function(response) {
                    if (!response.sucesso) {
                        console.error("Erro ao carregar atendentes: " + response.mensagem);
                        return;
                    }

                    response.dados.forEach(function(item) {
                        $select.append($('<option>', { value: item.id, text: item.nome }));
                    });
                },
                error: function(xhr, status, error) {
                    console.error('Erro ao carregar atendentes:', error);
                }
            });
        }

        function validarFiltros() {
            if (!$('#filtroDataCriacaoDe').val() || !$('#filtroDataCriacaoAte').val()) {
                Toast.error("Preencha a data de Criação");
                return false;
            }
            return true;
        }

        function exibirLoading(exibir) {
            carregando = exibir;
            if (exibir) {
                $('#loadingOverlay').removeClass('d-none').addClass('d-flex');
            } else {
                $('#loadingOverlay').removeClass('d-flex').addClass('d-none');
            }
        }

        function pesquisar() {
            if (!validarFiltros()) return;
            
            exibirLoading(true);

            const params = {
                acao: 'obterDadosDashboard',
                filtroFilial: $('#filtroFilial').val(),
                filtroDataCriacaoDe: $('#filtroDataCriacaoDe').val(),
                filtroDataCriacaoAte: $('#filtroDataCriacaoAte').val(),
                filtroDepartamento: $('#filtroDepartamento').val(),
                filtroCategoria: $('#filtroCategoria').val(),
                filtroStatus: $('#filtroStatus').val(),
                filtroPrioridade: $('#filtroPrioridade').val(),
                filtroAtendente: $('#filtroAtendente').val(),
                filtroSolicitante: $('#filtroSolicitante').val(),
            };

            $.ajax({
                url: 'DashboardRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: params,
                dataType: 'json',
                success: function(response) {
                    exibirLoading(false);

                    if (!response.sucesso) {
                        Toast.error(response.mensagem);
                        return;
                    }

                    dadosDashboard = response.dados || [];
                    controlarExibicaoAviso();
                    renderizarDashboard();
                },
                error: function(xhr, status, error) {
                    exibirLoading(false);
                    console.error('Erro ao pesquisar:', error);
                    Toast.error('Erro ao realizar a pesquisa');
                }
            });
        }

        function destruirCharts() {
            Object.keys(charts).forEach(key => {
                if (charts[key]) {
                    charts[key].destroy();
                    charts[key] = null;
                }
            });
        }

        function formatarDadosSimples(dados, horizontal) {
            const labels = dados.map(item => item.label);
            const totals = dados.map(item => item.total);

            datasets = [{
                label: 'Total',
                data: totals,
                backgroundColor: coresGrafico.slice(0, totals.length),
                borderColor: bordasGrafico.slice(0, totals.length),
                borderWidth: 0.55,
                borderRadius: 1,
                barThickness: horizontal ? 20 : 'flex',
                maxBarThickness: horizontal ? 24 : 50
            }];

            return { labels, datasets };
        }

        function formatarDadosAninhado(dados, horizontal) {
            const labels = dados.map(item => item.label);
            const seriesNomes = [...new Set(
                dados.flatMap(item => item.total.map(v => v.label))
            )];

            datasets = seriesNomes.map((serie, index) => ({
                label: serie,
                data: dados.map(item => {
                    const encontrado = item.total.find(v => v.label === serie);
                    return encontrado ? encontrado.total : 0;
                }),
                backgroundColor: coresGrafico[index % coresGrafico.length],
                borderColor: bordasGrafico[index % bordasGrafico.length],
                borderWidth: 0.55,
                borderRadius: 1,
                barThickness: horizontal ? 20 : 'flex',
                maxBarThickness: horizontal ? 24 : 50
            }));

            return { labels, datasets };
        }

        function criarGraficoBarras(canvasId, dados, horizontal = false, aninhado = false) {
            const ctx = document.getElementById(canvasId);
            if (!ctx) return null;

            const dadosFormatados = aninhado ? 
                formatarDadosAninhado(dados, horizontal) : formatarDadosSimples(dados, horizontal);

            console.log(dadosFormatados)

            const config = {
                type: 'bar',
                data: dadosFormatados,
                options: {
                    indexAxis: horizontal ? 'y' : 'x',
                    responsive: true,
                    maintainAspectRatio: false,
                    plugins: {
                        legend: { display: false },
                        tooltip: {
                            backgroundColor: '#212529',
                            titleFont: { size: 12 },
                            bodyFont: { size: 12 },
                            padding: 10,
                            cornerRadius: 2
                        }
                    },
                    scales: {
                        x: {
                            grid: {
                                display: !horizontal,
                                color: '#e9ecef'
                            },
                            ticks: {
                                font: { size: 11 },
                                color: '#6c757d'
                            }
                        },
                        y: {
                            grid: {
                                display: horizontal,
                                color: '#e9ecef'
                            },
                            ticks: {
                                font: { size: 11 },
                                color: '#6c757d'
                            },
                            beginAtZero: true
                        }
                    }
                }
            };

            return new Chart(ctx, config);
        }

        function criarGraficoPizza(canvasId, dados) {
            const ctx = document.getElementById(canvasId);
            if (!ctx) return null;

            const dadosFormatados = formatarDadosSimples(dados)
            const config = {
                type: 'doughnut',
                data: dadosFormatados,
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    plugins: {
                        legend: {
                            position: 'right',
                            labels: {
                                font: { size: 11 },
                                color: '#6c757d',
                                padding: 12,
                                usePointStyle: true,
                                pointStyle: 'circle'
                            }
                        },
                        tooltip: {
                            backgroundColor: '#212529',
                            titleFont: { size: 12 },
                            bodyFont: { size: 12 },
                            padding: 10,
                            cornerRadius: 4
                        }
                    },
                    cutout: '60%'
                }
            };

            return new Chart(ctx, config);
        }
        
        function alternarGraficoCategoria(tipo) {
            tipoGraficoCategoria = tipo;
            const ehCategoria = tipo === 'categoria';
            const dados = ehCategoria ? dadosDashboard.chamadoPorCategoria : dadosDashboard.chamadoPorTipo;

            $('#tituloGraficoCategoria').text(ehCategoria ? 'Chamados por categoria' : 'Chamados por tipo');
            $('#btnGraficoCategoria').toggleClass('active', ehCategoria);
            $('#btnGraficoTipo').toggleClass('active', !ehCategoria);

            if (charts.categoria) {
                charts.categoria.destroy();
                charts.categoria = null;
            }

            charts.categoria = criarGraficoPizza('chartCategoria', dados || []);
        }

        function criarGraficoPorcentagem(canvasId, dados) {
            const ctx = document.getElementById(canvasId);
            if (!ctx) return null;

            const dadosFiltrados = dados.filter(item => parseFloat(item.total) > 0);
            const avg = Math.round(dadosFiltrados.reduce((s, c) => s + parseFloat(c.total), 0) / dadosFiltrados.length);
            $('#avgLabel').text((avg || '0') + '%');

            $.each(dadosFiltrados, (_, cat) => {
                $('#catList').append(`
                <div>
                    <div class="d-flex justify-content-between mb-1">
                        <span class="text-secondary small">${cat.label}</span>
                        <span class="fw-semibold small">${formatarMoeda(cat.total)}%</span>
                    </div>
                    <div class="bar-bg">
                    <div class="bar-fill" style="width:${cat.total}%"></div>
                    </div>
                </div>`);
            });

            const config = {
                type: 'doughnut',
                data: {
                    datasets: [{
                        data: [avg, 100 - avg],
                        backgroundColor: ['rgba(25, 135, 84, 1)', '#efefef'],
                        borderWidth: 0.55,
                        borderRadius: 1,
                        barThickness: 20,
                        maxBarThickness: 24,
                    }]
                },
                options: {
                    cutout: '82%',
                    rotation: -90,
                    animation: { duration: 900, easing: 'easeInOutQuart' },
                    plugins: { legend: { display: false }, tooltip: { enabled: false } },
                    events: [],
                }
            };

            return new Chart(ctx, config);
        }

        function renderizarDashboard() {
            destruirCharts();

            if (!dadosDashboard) {
                $('#kpiAbertos').text('0');
                $('#kpiConcluidos').text('0');
                $('#kpiAtrasados').text('0');
                $('#kpiTempoMedio').text('0');
                
                $('#dashboardContainer').html(`
                    <div class="col-12">
                        <div class="card shadow-sm border-0">
                            <div class="card-body text-center py-5 text-muted">
                                <i class="bi bi-bar-chart-line display-4 mb-3 opacity-50"></i>
                                <p class="mb-0">Utilize os filtros acima e clique em Pesquisar para visualizar o dashboard</p>
                            </div>
                        </div>
                    </div>
                `);
                return;
            }

            $('#kpiAbertos').text(dadosDashboard.totalAbertos.toLocaleString('pt-BR'));
            $('#kpiConcluidos').text(dadosDashboard.totalConcluidos.toLocaleString('pt-BR'));
            $('#kpiAtrasados').text(dadosDashboard.totalAtrasados.toLocaleString('pt-BR'));
            $('#kpiTempoMedio').text(formatarMinutos(dadosDashboard.tempoMedioMinutos));

            $('#dashboardContainer').html(`
                <div class="col-12 col-lg-4">
                    <div class="card shadow-sm border-0 h-100">
                       <div class="d-flex justify-content-between border-bottom p-3 border-color-secondary">
                            <span class="h6 mb-0">
                                Chamados por Departamento
                            </span>
                        </div>
                        <div class="card-body">
                            <div class="chart-container">
                                <canvas id="chartDepartamento"></canvas>
                            </div>
                        </div>
                    </div>
                </div>

                <div class="col-12 col-lg-4">
                    <div class="card shadow-sm border-0 h-100">
                        <div class="d-flex justify-content-between align-items-center border-bottom p-3 border-color-secondary">
                            <span id="tituloGraficoCategoria" class="h6 mb-0">
                                Chamados por tipo de chamado
                            </span>
                            <div class="btn-group btn-group-sm" role="group" aria-label="Alternar gráfico de chamados">
                                <button type="button" id="btnGraficoCategoria" class="btn btn-outline-primary" onclick="alternarGraficoCategoria('categoria')">Categoria</button>
                                <button type="button" id="btnGraficoTipo" class="btn btn-outline-primary active" onclick="alternarGraficoCategoria('tipo')">Tipo</button>
                            </div>
                        </div>
                        <div class="card-body">
                            <div class="chart-container">
                                <canvas id="chartCategoria"></canvas>
                            </div>
                        </div>
                    </div>
                </div>

                <div class="col-12 col-lg-4">
                    <div class="card shadow-sm border-0 h-100">
                        <div class="d-flex justify-content-between border-bottom p-3 border-color-secondary">
                            <span class="h6 mb-0">
                                SLA cumprido por urgência
                            </span>
                        </div>
                        <div class="card-body">
                            <div class="d-flex gap-4 align-items-center justify-content-between p-3">
                                <div class="chart-wrap flex-shrink-0">
                                    <div class="chart-container">
                                        <canvas id="chartSLA"></canvas>
                                    </div>
                                    
                                    <div class="chart-center d-flex flex-column align-items-center justify-content-center">
                                        <span id="avgLabel" class="fw-bold" style="font-size:2rem"></span>
                                        <span class="text-uppercase text-secondary" style="font-size:.6rem;letter-spacing:.12em">Cumprido</span>
                                    </div>
                                </div>

                                <div id="catList" style="max-width: 220px" class="d-flex flex-column gap-3 w-100"></div>
                            </div>
                        </div>
                    </div>
                </div>

                <div class="col-12 col-lg-7">
                    <div class="card shadow-sm border-0 h-100">
                       <div class="d-flex justify-content-between border-bottom p-3 border-color-secondary">
                            <span class="h6 mb-0">
                                Volume por mês <span class="text-muted !fs-6 fw-normal">(Últimos 12 meses)</span>
                            </span>
                        </div>
                        <div class="card-body">
                            <div class="chart-container">
                                <canvas id="chartVolume"></canvas>
                            </div>
                        </div>
                    </div>
                </div>

                <div class="col-12 col-lg-5">
                    <div class="card shadow-sm border-0 h-100">
                        <div class="d-flex justify-content-between border-bottom p-3 border-color-secondary">
                            <span class="h6 mb-0">
                                Performance da equipe
                            </span>
                        </div>
                        <div class="card-body">
                            <table class="table table-hover mb-0" id="gradeAtendente">
                                <thead>
                                    <tr>
                                        <th>Atendente</th>
                                        <th>Resolvidos</th>
                                        <th>Resolvidos no Prazo</th>
                                        <th>Percentual SLA</th>
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
            `);


            alternarGraficoCategoria(tipoGraficoCategoria);
            charts.departamento = criarGraficoBarras('chartDepartamento', dadosDashboard.chamadoPorDepartamento, true);
            charts.volume = criarGraficoBarras('chartVolume', dadosDashboard.volumeMensal, false, true);
            charts.porcentagemSLA = criarGraficoPorcentagem('chartSLA', dadosDashboard.SLAPorUrgencia, true);
            renderizarGradeAtendente(dadosDashboard.chamadoPorAtendente);
        }

        function renderizarGradeAtendente(dados){
            const $tbody = $('#gradeAtendente tbody');
            
            $tbody.empty();
            
            if (dados.length === 0) {
                $tbody.html(`
                    <tr>
                        <td colspan="28" class="text-center text-muted py-4">
                            Nenhum registro encontrado
                        </td>
                    </tr>
                `);
                return;
            }

            dados.forEach(function(item) {   
                const percentualCumprido = parseInt(item.quantidadeResolvidaNoPrazo / (parseInt(item.quantidadeResolvida) || 1)) * 100;
 
                $tbody.append(`
                    <tr>
                        <td>${item.nome || '-'}</td>
                        <td>${item.quantidadeResolvida || '0'}</td>
                        <td>${item.quantidadeResolvidaNoPrazo || '0'}</td>
                        <td>${percentualCumprido || '0'}%</td>
                    </tr>
                `);
            });
        }

        function abrirGraficoCompleto(tipoGrafico) {  
            let url = `DashboardFNCDetalhes.asp?tipoGrafico=${tipoGrafico}`;

            if ($('#filtroFilial').val()) url += `&filtroFilial=${$('#filtroFilial').val()}`;
            if ($('#filtroDepartamento').val()) url += `&filtroDepartamento=${$('#filtroDepartamento').val()}`;
            if ($('#filtroDataCriacaoDe').val()) url += `&filtroDataCriacaoDe=${$('#filtroDataCriacaoDe').val()}`;
            if ($('#filtroDataCriacaoAte').val()) url += `&filtroDataCriacaoAte=${$('#filtroDataCriacaoAte').val()}`;
            if ($('#filtroCategoria').val()) url += `&filtroCategoria=${$('#filtroCategoria').val()}`;

            window.open(url, '_blank');
        }

        function abrirChamadosAtrasados() {
            let url = `../Chamado/ChamadoGrade.asp?filtroStatus=FINALIZADO_COM_ATRASO`;
            if ($('#filtroFilial').val()) url += `&filtroFilial=${$('#filtroFilial').val()}`;
            if ($('#filtroDepartamento').val()) url += `&filtroDepartamento=${$('#filtroDepartamento').val()}`;
            if ($('#filtroDataCriacaoDe').val()) url += `&filtroDataCriacaoDe=${$('#filtroDataCriacaoDe').val()}`;
            if ($('#filtroDataCriacaoAte').val()) url += `&filtroDataCriacaoAte=${$('#filtroDataCriacaoAte').val()}`;
            if ($('#filtroCategoria').val()) url += `&filtroCategoria=${$('#filtroCategoria').val()}`;
            if ($('#filtroDepartamento').val()) url += `&filtroDepartamento=${$('#filtroDepartamento').val()}`;
            if ($('#filtroPrioridade').val()) url += `&filtroPrioridade=${$('#filtroPrioridade').val()}`;
            if ($('#filtroSolicitante').val()) url += `&filtroSolicitante=${$('#filtroSolicitante').val()}`;
            if ($('#filtroAtendente').val()) url += `&filtroAtendente=${$('#filtroAtendente').val()}`;

            window.open(url, '_blank');
        }

        function limparFiltros() {
            $('#filtroFilial').val('').trigger('change');
            $('#filtroDepartamento').val('').trigger('change');
            $('#filtroDataCriacaoDe').val('');
            $('#filtroDataCriacaoAte').val('');
            $('#filtroCategoria').val('').trigger('change');
            $('#filtroStatus').val('').trigger('change');
            $('#filtroPrioridade').val('').trigger('change');
            $('#filtroAtendente').val('').trigger('change');
            $('#filtroSolicitante').val('');
            
            dadosDashboard = null;
            renderizarDashboard();
        }

        function anexarEventos() {
            $('#filtroFilial').on('change', function() {
                campoPlantaAlterado();
            });

            $('#filtroDepartamento').on('change', function() {
                campoDepartamentoAlterado();
            });

            $('#btnPesquisar').on('click', pesquisar);
            $('#btnLimpar').on('click', limparFiltros);
        }

        function formatarMinutos(totalMinutos) {
            if (totalMinutos < 0) throw new Error("O valor não pode ser negativo.");
            if (totalMinutos === 0) return "0 minutos";

            const horas = Math.floor(totalMinutos / 60);
            const minutos = totalMinutos % 60;

            const partes = [];
            if (horas > 0)   partes.push(`${horas} h`);
            if (minutos > 0) partes.push(`${minutos} min`);

            return partes.join(" e ");
        }

        function controlarExibicaoAviso() {    
            if (!$('#filtroFilial').val()) {
                $('#avisoFilial').removeClass('d-none');
            } else {
                $('#avisoFilial').addClass('d-none');
            }
        }

        $(document).ready(function() {            
            inicializarSelect2();
            carregarDepartamentos();
            carregarCategorias();
            anexarEventos();
            renderizarDashboard();
        });
    </script>
</head>
<body>
    <div id="loadingOverlay" class="position-fixed top-0 start-0 w-100 h-100 d-none justify-content-center align-items-center" style="background: rgba(255,255,255,0.8); z-index: 9999;">
        <div class="spinner-border text-primary" role="status">
            <span class="visually-hidden">Carregando...</span>
        </div>
    </div>

    <div class="container my-4">
        <div class="d-flex my-4">
            <h3 class="mb-0">Dashboard de Chamados</h3>
        </div>

        <div class="card shadow-sm border-0 mb-4">
            <div class="card-body">
                <div class="row g-2">
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
                        <label for="filtroDataCriacaoDe" class="form-label">Data Criação De</label>
                        <input type="date" class="form-control form-control-sm" id="filtroDataCriacaoDe" name="filtroDataCriacaoDe">
                    </div>

                    <div class="col-12 col-md-4 col-lg-3">
                        <label for="filtroDataCriacaoAte" class="form-label">Data Criação Até</label>
                        <input type="date" class="form-control form-control-sm" id="filtroDataCriacaoAte" name="filtroDataCriacaoAte">
                    </div>

                    <div class="col-12 col-md-4 col-lg-3">
                        <label for="filtroAtendente" class="form-label">Atendente</label>
                        <select disabled class="form-select form-select-sm select2" id="filtroAtendente" name="filtroAtendente" data-placeholder="Todos">
                            <option value="">Todos</option>
                        </select>
                    </div>

                    <div class="col-12 col-md-4 col-lg-3">
                        <label for="filtroDepartamento" class="form-label">Departamento</label>
                        <select class="form-select form-select-sm select2" id="filtroDepartamento" name="filtroDepartamento" data-placeholder="Todos os departamentos">
                            <option value="">Todos os departamentos</option>
                        </select>
                    </div>

                    <div class="col-12 col-md-4 col-lg-3">
                        <label for="filtroCategoria" class="form-label">Categoria</label>
                        <select class="form-select form-select-sm select2" id="filtroCategoria" name="filtroCategoria" data-placeholder="Todos as categorias">
                            <option value="">Todos as categorias</option>
                        </select>
                    </div>

                    <div class="col-12 col-md-4 col-lg-3">
                        <label for="filtroStatus" class="form-label">Status</label>
                        <select class="form-select form-select-sm select2" id="filtroStatus" name="filtroStatus" data-placeholder="Todos">
                            <option value="">Todos</option>
                            <option value="PENDENTE">Pendente</option>
                            <option value="ABERTO">Aberto</option>
                            <option value="EM_ANALISE">Em Análise</option>
                            <option value="AGUARDANDO_TERCEIROS">Aguardando Terceiros</option>
                            <option value="FINALIZADO">Finalizado</option>
                            <option value="FINALIZADO_COM_ATRASO">Finalizado com Atraso</option>
                            <option value="CANCELADO">Cancelado</option>
                        </select>
                    </div>

                    <div class="col-12 col-md-4 col-lg-3">
                        <label for="filtroPrioridade" class="form-label">Prioridade</label>
                        <select class="form-select form-select-sm select2" id="filtroPrioridade" name="filtroPrioridade" data-placeholder="Todos">
                            <option value="">Todos</option>
                            <option value="BAIXA">Baixa</option>
                            <option value="MEDIA">Média</option>
                            <option value="ALTA">Alta</option>
                            <option value="URGENTE">Urgente</option>
                        </select>
                    </div>

                    <div class="col-12 col-md-4 col-lg-3">
                        <label for="filtroSolicitante" class="form-label">Solicitante</label>
                        <input placeholder="Email ou nome" type="text" class="form-control form-control-sm" id="filtroSolicitante" name="filtroSolicitante">
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
        </div>

        <div id="avisoFilial" class="d-none alert alert-warning mb-4">
            Atenção! Você está visualizando dados de todas as filiais. Para ver dados de uma filial específica, selecione-a no filtro acima.
        </div>

        <div class="row g-3 mb-4">
            <div class="col-6 col-lg-3">
                <div class="card shadow-sm border-0 h-100">
                    <div class="card-body d-flex justify-content-between align-items-center">
                        <div>
                            <div class="text-secondary mb-1">Abertos</div>
                            <h2 id="kpiAbertos">0</h2>
                        </div>
                        <div class="icon-wrapper purple">
                            <i class="bi bi-file-earmark-text"></i>
                        </div>
                    </div>
                </div>
            </div>
            <div class="col-6 col-lg-3">
                <div class="card shadow-sm border-0 h-100 border-start">
                    <div class="card-body d-flex justify-content-between align-items-center">
                        <div>
                            <div class="text-secondary mb-1">Resolvidos</div>
                            <h2 id="kpiConcluidos">0</h2>
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
                            <div class="text-secondary mb-1">Resolvidos com Atraso</div>
                            <h2 id="kpiAtrasados">0</h2>
                        </div>
                        <div class="icon-wrapper red">
                            <i class="bi bi-exclamation-triangle"></i>
                        </div>
                    </div>
                </div>
            </div>
            <div class="col-6 col-lg-3">
                <div class="card shadow-sm border-0 h-100 border-start">
                    <div class="card-body d-flex justify-content-between align-items-center">
                        <div>
                            <div class="text-secondary mb-1">Tempo Médio de Atendimento</div>
                            <h2 id="kpiTempoMedio">0</h2>
                        </div>
                        <div class="icon-wrapper yellow">
                            <i class="bi bi-clock"></i>
                        </div>
                    </div>
                </div>
            </div>
        </div>

        <div class="row g-3" id="dashboardContainer"></div>
    </div>
</body>
</html>